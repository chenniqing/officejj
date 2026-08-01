package cn.javaex.officejj.pdf.help;

import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.Map;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfStamper;

import cn.javaex.officejj.common.entity.Font;
import cn.javaex.officejj.common.entity.Picture;
import cn.javaex.officejj.common.entity.RGB;

/**
 * 表单填充
 *
 * @author 陈霓清
 */
public class AcroFieldsHelp extends Helper {
	/**
	 * 替换占位符内容
	 * @param form
	 * @param stamper
	 * @param param
	 * @return
	 * @throws IOException
	 * @throws DocumentException
	 */
	public AcroFields replaceContent(AcroFields form, PdfStamper stamper, Map<String, Object> param) throws IOException, DocumentException {
		return replaceContent(form, stamper, param, null);
	}

	/**
	 * 替换占位符内容。
	 * 支持传入默认字体，适合中文业务表单统一指定 SimSun、微软雅黑等字体，避免 PDF 表单扁平化后中文不显示。
	 * @param form PDF表单对象
	 * @param stamper PDF写入器
	 * @param param 字段名和值，字段名必须和PDF模板中的表单域名称一致
	 * @param defaultFontFamily 默认字体路径，支持绝对路径、相对路径和 resources: 前缀
	 * @return
	 * @throws IOException
	 * @throws DocumentException
	 */
	public AcroFields replaceContent(AcroFields form, PdfStamper stamper, Map<String, Object> param, String defaultFontFamily) throws IOException, DocumentException {
		if (param==null || param.size()==0) {
			return form;
		}

		BaseFont defaultBaseFont = this.createBaseFont(defaultFontFamily);
		for (Map.Entry<String, Object> entry : param.entrySet()) {
			String key = entry.getKey();
			Object value = entry.getValue();

			if (value==null) {
				value = "";
			}

			// 文本替换
			if (value instanceof String) {
				this.applyDefaultFont(form, key, defaultBaseFont);
				form.setField(key, (String) value);
			}
			// 自定义字体样式
			else if (value instanceof Font) {
				Font font = (Font) value;

				if (font.getFontFamily()!=null) {
					String path = super.getRealPath(font.getFontFamily());
					BaseFont baseFont = BaseFont.createFont(path, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
					form.setFieldProperty(key, "textfont", baseFont, null);
				} else {
					this.applyDefaultFont(form, key, defaultBaseFont);
				}
				if (font.getColor()!=null) {
					RGB rgb = new RGB(font.getColor());
					BaseColor baseColor = new BaseColor(rgb.getRed(), rgb.getGreen(), rgb.getBlue());
					form.setFieldProperty(key, "textcolor", baseColor, null);
				}
				if (font.getFontSize()!=null) {
					form.setFieldProperty(key, "textsize", font.getFontSize().floatValue(), null);
				}

				form.setField(key, font.getText());
			}
			// 图片替换
			else if (value instanceof Picture) {
				Picture picture = (Picture) value;
				if (form.getFieldPositions(key)==null || form.getFieldPositions(key).isEmpty()) {
					continue;
				}

				// 获取所在页和坐标，左下角为起点
				// 图片字段只借用 PDF 表单域坐标；真正输出内容由图片覆盖。
				// 可视化模板可能把字段默认值设置为 image，先清空字段值，避免扁平化后把占位文字一并输出。
				form.setField(key, "");
				int pageNo = form.getFieldPositions(key).get(0).page;
				Rectangle signRect = form.getFieldPositions(key).get(0).position;
				float x = signRect.getLeft();
				float y = signRect.getBottom();

				// 读取图片。URL中可能包含中文文件名，iText直接打开时不会兜底编码，文件服务可能返回400。
				Image image = this.createImage(picture);
				// 获取操作的页面
				PdfContentByte under = stamper.getOverContent(pageNo);
				// 设置图片大小
				if (picture.getWidth()==null || picture.getHeight()==null) {
					image.scaleToFit(signRect.getWidth(), signRect.getHeight());    // 根据域的大小缩放图片（图片大小自适应）
				} else {
					double width = picture.getWidth();
					double height = picture.getHeight();
					image.scaleAbsolute((float) width, (float) height);    // 指定图片大小
				}
				// 添加图片
				image.setAbsolutePosition(x, y);
				under.addImage(image);
			}
			// 数字之类的直接转字符串
			else {
				this.applyDefaultFont(form, key, defaultBaseFont);
				form.setField(key, value.toString());
			}
		}

		return form;
	}

	/**
	 * 创建默认字体对象。
	 * 字体为空时沿用PDF模板字段自身的外观设置，保持旧版本行为不变。
	 */
	private BaseFont createBaseFont(String fontFamily) throws DocumentException, IOException {
		if (fontFamily==null || fontFamily.trim().length()==0) {
			return null;
		}

		String path = super.getRealPath(fontFamily);
		return BaseFont.createFont(path, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
	}

	/**
	 * 创建图片对象。
	 * 优先使用调用方传入的字节数组；只有传入URL时才走网络加载，并在原始URL失败后尝试百分号编码后的URL。
	 */
	private Image createImage(Picture picture) throws IOException, DocumentException {
		if (picture.getData()!=null) {
			return Image.getInstance(picture.getData());
		}
		String url = picture.getUrl();
		if (url==null || url.trim().length()==0) {
			throw new IOException("图片地址不能为空");
		}

		String encodedUrl = this.encodeUrl(url);
		try {
			return Image.getInstance(encodedUrl);
		} catch (IOException e) {
			if (encodedUrl.equals(url)) {
				throw e;
			}
			try {
				return Image.getInstance(url);
			} catch (IOException retryException) {
				retryException.addSuppressed(e);
				throw retryException;
			}
		}
	}

	/**
	 * 把URL转换为ASCII形式，主要处理query/path中的中文、空格等字符。
	 */
	private String encodeUrl(String url) throws MalformedURLException {
		try {
			return new URI(url).toASCIIString();
		} catch (URISyntaxException e) {
			URL parsedUrl = new URL(url);
			try {
				return new URI(
						parsedUrl.getProtocol(),
						parsedUrl.getUserInfo(),
						parsedUrl.getHost(),
						parsedUrl.getPort(),
						parsedUrl.getPath(),
						parsedUrl.getQuery(),
						parsedUrl.getRef()
				).toASCIIString();
			} catch (URISyntaxException uriException) {
				MalformedURLException malformedURLException = new MalformedURLException("图片URL编码失败：" + url);
				malformedURLException.initCause(uriException);
				throw malformedURLException;
			}
		}
	}

	/**
	 * 给普通文本字段套用默认字体。
	 * 如果模板中不存在该字段，iText会返回false；这里不抛错，允许同一份参数Map复用于不同模板。
	 */
	private void applyDefaultFont(AcroFields form, String key, BaseFont defaultBaseFont) {
		if (defaultBaseFont==null) {
			return;
		}
		form.setFieldProperty(key, "textfont", defaultBaseFont, null);
	}

}
