package cn.javaex.officejj.pdf.help;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PRStream;
import com.itextpdf.text.pdf.PRTokeniser;
import com.itextpdf.text.pdf.PdfBoolean;
import com.itextpdf.text.pdf.PdfContentParser;
import com.itextpdf.text.pdf.PdfDictionary;
import com.itextpdf.text.pdf.PdfName;
import com.itextpdf.text.pdf.PdfNumber;
import com.itextpdf.text.pdf.PdfObject;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.RandomAccessFileOrArray;
import com.itextpdf.text.pdf.parser.ImageRenderInfo;
import com.itextpdf.text.pdf.parser.PdfContentStreamProcessor;
import com.itextpdf.text.pdf.parser.RenderListener;
import com.itextpdf.text.pdf.parser.TextRenderInfo;
import com.itextpdf.text.pdf.parser.Vector;

/**
 * 处理模板显式启用的文字排版，不改变普通 PDF 字段的填充行为。
 * OfficeJJVerticalAlign=Center 启用垂直居中，OfficeJJJoinLines=true 将输入换行转换为间隔。
 * 使用 iText 原生换行结果测量文字，保留字体、颜色、边框和表单值。
 */
public class TextFieldLayoutHelper {

	public static final PdfName VERTICAL_ALIGN = new PdfName("OfficeJJVerticalAlign");
	public static final PdfName JOIN_LINES = new PdfName("OfficeJJJoinLines");
	private final Map<String, BaseFont> centeredFonts = new LinkedHashMap<String, BaseFont>();

	/**
	 * 按模板配置准备显示文本，并记录需要后续居中的字段。
	 *
	 * @param form 当前表单
	 * @param key 字段名称
	 * @param text 待填充文字，null 按空字符串处理
	 * @param baseFont 调用方指定的字体，为 null 时使用 PDF 自带字体度量
	 * @return 按配置处理后的显示文字，不修改调用方传入的对象
	 */
	public String prepareText(AcroFields form, String key, String text, BaseFont baseFont) {
		String value = text == null ? "" : text;
		AcroFields.Item item = form.getFieldItem(key);
		if (item == null || form.getFieldType(key) != AcroFields.FIELD_TYPE_TEXT) {
			return value;
		}
		if (PdfBoolean.PDFTRUE.equals(item.getMerged(0).getAsBoolean(JOIN_LINES))) {
			value = value.replaceAll("\\r\\n|[\\r\\n\\u0085\\u2028\\u2029]", "  ").trim();
		}
		if (!value.trim().isEmpty()) {
			for (int index = 0; index < item.size(); index++) {
				if (PdfName.CENTER.equals(item.getMerged(index).getAsName(VERTICAL_ALIGN))) {
					centeredFonts.put(key, baseFont);
					break;
				}
			}
		}
		return value;
	}

	/**
	 * 判断本次填充是否需要在扁平化之前调整文字外观。
	 *
	 * @return 仅存在已填充且启用了垂直居中的文字字段时返回 true
	 */
	public boolean hasCenteredFields() {
		return !centeredFonts.isEmpty();
	}

	/**
	 * 调整已生成的字段外观，再按调用方要求保留表单或扁平化。
	 * 先完成字体和外观资源写入，再进行测量，避免依赖 iText 尚未落盘的内部对象。
	 *
	 * @param filledPdf 已填充但尚未扁平化的 PDF
	 * @param readOnly 是否将最终结果扁平化
	 * @return 排版后的 PDF 输出流
	 * @throws IOException 字段外观异常或文字无法完整容纳时抛出，包含字段名称
	 * @throws DocumentException PDF 写入失败时抛出
	 */
	public ByteArrayOutputStream apply(ByteArrayOutputStream filledPdf, boolean readOnly) throws IOException, DocumentException {
		PdfReader reader = new PdfReader(filledPdf.toByteArray());
		try {
			ByteArrayOutputStream output = new ByteArrayOutputStream();
			PdfStamper stamper = new PdfStamper(reader, output);
			AcroFields form = stamper.getAcroFields();
			TextBounds textBounds = new TextBounds();
			for (Map.Entry<String, BaseFont> entry : centeredFonts.entrySet()) {
				AcroFields.Item item = form.getFieldItem(entry.getKey());
				for (int index = 0; index < item.size(); index++) {
					if (PdfName.CENTER.equals(item.getMerged(index).getAsName(VERTICAL_ALIGN))) {
						centerAppearance(stamper, item.getMerged(index), entry.getKey(), entry.getValue(), textBounds);
					}
				}
			}
			stamper.setFormFlattening(readOnly);
			stamper.close();
			return output;
		} finally {
			// 输出仅存于内存；发生异常时丢弃未完成结果，始终释放读取器。
			reader.close();
		}
	}

	/**
	 * 将一个字段中的文字移动到可用区域中央，保持边框及裁剪区域原位。
	 *
	 * @param stamper 当前写入器
	 * @param field 包含继承属性的字段字典
	 * @param key 用于异常定位的字段名称
	 * @param baseFont 调用方指定的字体，可为 null
	 * @param textBounds 可复用的文字测量器
	 * @throws IOException 外观不受支持或文字超出可用区域时抛出
	 */
	private void centerAppearance(PdfStamper stamper, PdfDictionary field, String key, BaseFont baseFont,
			TextBounds textBounds) throws IOException {
		PdfDictionary appearanceDictionary = field.getAsDict(PdfName.AP);
		PdfObject normal = appearanceDictionary == null ? null : PdfReader.getPdfObject(appearanceDictionary.get(PdfName.N));
		if (!(normal instanceof PRStream)) {
			throw new IOException("PDF字段缺少文字外观：" + key);
		}
		PdfDictionary characteristics = field.getAsDict(PdfName.MK);
		PdfNumber rotation = characteristics == null ? null : characteristics.getAsNumber(PdfName.R);
		if (rotation != null && rotation.intValue() % 360 != 0) {
			throw new IOException("垂直居中字段暂不支持旋转：" + key);
		}
		PRStream appearance = (PRStream) normal;
		Rectangle box = PdfReader.getNormalizedRectangle(appearance.getAsArray(PdfName.BBOX));
		PdfDictionary border = field.getAsDict(PdfName.BS);
		float margin = border != null && border.getAsNumber(PdfName.W) != null ? border.getAsNumber(PdfName.W).floatValue() : 0;
		if (border != null && (PdfName.B.equals(border.getAsName(PdfName.S)) || PdfName.I.equals(border.getAsName(PdfName.S)))) {
			margin *= 2;
		}
		byte[] content = PdfReader.getStreamBytes(appearance);
		textBounds.measure(content, appearance.getAsDict(PdfName.RESOURCES), baseFont);
		if (field.getAsString(PdfName.V) != null
				&& textBounds.characterCount < countVisibleCharacters(field.getAsString(PdfName.V).toUnicodeString())) {
			throw new IOException("PDF字段文字未完整排版，请扩大模板字段尺寸：" + key);
		}
		if (!textBounds.hasText) {
			return;
		}
		if (textBounds.right > box.getRight() - margin + 0.05f || textBounds.left < box.getLeft() + margin - 0.05f
				|| textBounds.top - textBounds.bottom > box.getHeight() - 2 * margin + 0.05f) {
			throw new IOException("PDF字段文字超出可用区域，请扩大模板字段尺寸：" + key
					+ "，文字范围=[" + textBounds.left + "," + textBounds.bottom + "," + textBounds.right + "," + textBounds.top
					+ "]，字段尺寸=" + box.getWidth() + "x" + box.getHeight());
		}
		float offset = (box.getTop() + box.getBottom() - textBounds.top - textBounds.bottom) / 2;
		appearance.setData(translateText(content, offset, key));
		stamper.markUsed(appearance);
	}

	/**
	 * 统计非空白字符，检查 iText 是否因字段高度不足而省略了后续行。
	 *
	 * @param text 原始或已绘制的文本
	 * @return 非空白 Unicode 字符数量
	 */
	private static long countVisibleCharacters(String text) {
		return text.codePoints().filter(codePoint -> !Character.isWhitespace(codePoint)).count();
	}

	/**
	 * 使用 PDF 语法解析器调整文字块，不移动背景、边框或原有裁剪路径。
	 *
	 * @param content iText 生成的字段外观内容
	 * @param offset 垂直平移量，单位为点
	 * @param key 用于异常定位的字段名称
	 * @return 更新后的外观内容
	 * @throws IOException 内容解析失败或缺少标准文字块时抛出
	 */
	private byte[] translateText(byte[] content, float offset, String key) throws IOException {
		PRTokeniser tokeniser = new PRTokeniser(new RandomAccessFileOrArray(content));
		try {
			PdfContentParser parser = new PdfContentParser(tokeniser);
			ArrayList<PdfObject> operands = new ArrayList<PdfObject>();
			ByteArrayOutputStream output = new ByteArrayOutputStream();
			int depth = 0;
			int variableDepth = -1;
			boolean translated = false;
			while (!parser.parse(operands).isEmpty()) {
				String operator = operands.get(operands.size() - 1).toString();
				if ("BMC".equals(operator) || "BDC".equals(operator)) {
					depth++;
					if (PdfName.TX.equals(operands.get(0))) {
						variableDepth = depth;
					}
				}
				if (variableDepth > 0 && "BT".equals(operator)) {
					output.write(("q 1 0 0 1 0 " + new PdfNumber(offset) + " cm\n").getBytes(StandardCharsets.US_ASCII));
					translated = true;
				}
				for (PdfObject operand : operands) {
					operand.toPdf(null, output);
					output.write(' ');
				}
				output.write('\n');
				if (variableDepth > 0 && "ET".equals(operator)) {
					output.write("Q\n".getBytes(StandardCharsets.US_ASCII));
				}
				if ("EMC".equals(operator)) {
					if (variableDepth == depth) {
						variableDepth = -1;
					}
					depth--;
				}
			}
			if (!translated) {
				throw new IOException("PDF字段缺少可调整的文字块：" + key);
			}
			return output.toByteArray();
		} finally {
			tokeniser.close();
		}
	}

	/**
	 * 按实际字体和字号测量文字边界，复用解析器的字体缓存以支持批量字段。
	 */
	private static class TextBounds implements RenderListener {
		private final PdfContentStreamProcessor pdfContentStreamProcessor = new PdfContentStreamProcessor(this);
		private BaseFont baseFont;
		private boolean hasText;
		private long characterCount;
		private float left;
		private float right;
		private float bottom;
		private float top;

		/**
		 * 清空上一个字段的度量并解析当前外观。
		 *
		 * @param content 字段外观内容
		 * @param resources 外观引用的字体等资源
		 * @param baseFont 调用方指定的字体，可为 null
		 */
		private void measure(byte[] content, PdfDictionary resources, BaseFont baseFont) {
			this.baseFont = baseFont;
			hasText = false;
			characterCount = 0;
			left = bottom = Float.POSITIVE_INFINITY;
			right = top = Float.NEGATIVE_INFINITY;
			pdfContentStreamProcessor.reset();
			pdfContentStreamProcessor.processContent(content, resources);
		}

		/**
		 * 根据已排版字符的位置、字体和实际字号计算可见边界。
		 *
		 * @param info 当前绘制的文字信息
		 */
		@Override
		public void renderText(TextRenderInfo info) {
			characterCount += countVisibleCharacters(info.getText());
			for (TextRenderInfo character : info.getCharacterRenderInfos()) {
				String text = character.getText();
				if (text.trim().isEmpty()) {
					continue;
				}
				hasText = true;
				Vector origin = character.getBaseline().getStartPoint();
				String actualFont = character.getFont().getPostscriptFontName();
				boolean sameFont = baseFont != null && (actualFont.equals(baseFont.getPostscriptFontName())
						|| actualFont.endsWith("+" + baseFont.getPostscriptFontName()));
				int[] glyph = sameFont ? baseFont.getCharBBox(text.codePointAt(0)) : null;
				if (glyph != null) {
					float scale = pdfContentStreamProcessor.gs().getFontSize() / 1000;
					left = Math.min(left, origin.get(Vector.I1) + glyph[0] * scale);
					right = Math.max(right, origin.get(Vector.I1) + glyph[2] * scale);
					bottom = Math.min(bottom, origin.get(Vector.I2) + glyph[1] * scale);
					top = Math.max(top, origin.get(Vector.I2) + glyph[3] * scale);
				} else {
					left = Math.min(left, character.getDescentLine().getStartPoint().get(Vector.I1));
					right = Math.max(right, character.getAscentLine().getEndPoint().get(Vector.I1));
					bottom = Math.min(bottom, character.getDescentLine().getStartPoint().get(Vector.I2));
					top = Math.max(top, character.getAscentLine().getStartPoint().get(Vector.I2));
				}
			}
		}

		/**
		 * 文本块开始时沿用当前字段的累计边界。
		 */
		@Override
		public void beginTextBlock() {
		}

		/**
		 * 文本块结束时保留已测得的边界。
		 */
		@Override
		public void endTextBlock() {
		}

		/**
		 * 图片不参与文字的垂直对齐计算。
		 *
		 * @param info 图片绘制信息
		 */
		@Override
		public void renderImage(ImageRenderInfo info) {
		}
	}
}
