package cn.javaex.officejj.pdf;

import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.security.PrivateKey;
import java.security.cert.Certificate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.util.IOUtils;

import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfSignatureAppearance;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.security.BouncyCastleDigest;
import com.itextpdf.text.pdf.security.DigestAlgorithms;
import com.itextpdf.text.pdf.security.ExternalDigest;
import com.itextpdf.text.pdf.security.ExternalSignature;
import com.itextpdf.text.pdf.security.MakeSignature;
import com.itextpdf.text.pdf.security.PrivateKeySignature;
import com.itextpdf.tool.xml.XMLWorkerHelper;

import cn.javaex.officejj.common.util.PathHandler;
import cn.javaex.officejj.pdf.help.AcroFieldsHelp;
import cn.javaex.officejj.pdf.help.FontFamilyHelper;
import cn.javaex.officejj.pdf.help.MergeHelper;
import cn.javaex.officejj.pdf.help.PdfHelper;

/**
 * PDF工具类
 *
 * @author 陈霓清
 */
public class PdfUtils {

	/**
	 * 通过路径读取Pdf
	 * @param filePath     例如：D:\\123.pdf
	 * @return
	 * @throws IOException
	 */
	public static PdfReader getPdf(String filePath) throws IOException {
		return new PdfReader(filePath);
	}

	/**
	 * 读取resources文件夹下的Pdf
	 * @param filePath      resources文件夹下的路径，例如：template/pdf/模板.pdf
	 * @return
	 * @throws IOException
	 */
	public static PdfReader getPdfFromResource(String filePath) throws IOException {
		InputStream in = PathHandler.getInputStreamFromResource(filePath);
		return getPdf(in, true);
	}

	/**
	 * 通过流读取Pdf，不关闭调用方传入的输入流。
	 * @param in
	 * @return
	 */
	public static PdfReader getPdf(InputStream in) {
		return getPdf(in, false);
	}

	/**
	 * 通过流读取Pdf。
	 * 工具类内部创建的输入流需要关闭；调用方传入的输入流默认由调用方自己关闭。
	 * @param in
	 * @param closeInputStream 是否在读取完成后关闭输入流
	 * @return
	 */
	private static PdfReader getPdf(InputStream in, boolean closeInputStream) {
		if (in==null) {
			throw new IllegalArgumentException("PDF输入流不能为空");
		}

		try {
			return new PdfReader(in);
		} catch (Exception e) {
			throw new RuntimeException("读取PDF失败", e);
		} finally {
			if (closeInputStream) {
				IOUtils.closeQuietly(in);
			}
		}
	}

	/**
	 * 替换Pdf中占位符的内容（只读）
	 * @param reader
	 * @param param
	 * @return
	 * @throws IOException
	 */
	public static ByteArrayOutputStream writePdf(PdfReader reader, Map<String, Object> param) {
		return writePdf(reader, param, true);
	}

	/**
	 * 替换Pdf中占位符的内容
	 * @param reader
	 * @param param
	 * @param readOnly      是否只读
	 * @return
	 */
	public static ByteArrayOutputStream writePdf(PdfReader reader, Map<String, Object> param, boolean readOnly) {
		AcroFieldsHelp acroFieldsHelp = new AcroFieldsHelp();

		ByteArrayOutputStream bos = null;
		PdfStamper stamper = null;

		try {
			bos = new ByteArrayOutputStream();
			stamper = new PdfStamper(reader, bos);
			AcroFields form = stamper.getAcroFields();

			// 替换占位符为数据中的内容
			form = acroFieldsHelp.replaceContent(form, stamper, param);

			stamper.setFormFlattening(readOnly);    // 如果为false那么生成的PDF文件还能编辑
			stamper.close();
			stamper = null;

			reader.close();
		} catch (Exception e) {
			throw new RuntimeException("写入PDF表单失败", e);
		} finally {
			if (stamper!=null) {
				try { stamper.close(); } catch (Exception ignore) {}
			}
			if (reader!=null) {
				reader.close();
			}
		}

		return bos;
	}

	/**
	 * 输出Pdf到指定路径
	 * @param word
	 * @param filePath     文件写到哪里的全路径，例如：D:\\1.pdf
	 * @throws IOException
	 */
	public static void output(ByteArrayOutputStream bos, String filePath) throws IOException {
		// 保证这个文件的父文件夹必须要存在
		File targetFile = new File(filePath);
		File parentFile = targetFile.getParentFile();
		if (parentFile!=null && !parentFile.exists()) {
			parentFile.mkdirs();
		}
		OutputStream out = null;
		try {
			out = new FileOutputStream(targetFile);
			out.write(bos.toByteArray());
			out.flush();
		} catch (Exception e) {
			throw new IOException("输出PDF失败：" + filePath, e);
		} finally {
			IOUtils.closeQuietly(bos);
			IOUtils.closeQuietly(out);
		}
	}

	/**
	 * 下载Pdf（兼容 javax 和 jakarta Servlet 环境）
	 * @param response    javax.servlet.http.HttpServletResponse 或 jakarta.servlet.http.HttpServletResponse
	 * @param bos         PDF内容（ByteArrayOutputStream）
	 * @param filename    带后缀的文件名，例如："test.pdf"
	 */
	public static void download(Object response, ByteArrayOutputStream bos, String filename) throws IOException {
		OutputStream out = null;
		try {
			Method setContentType = response.getClass().getMethod("setContentType", String.class);
			Method setHeader = response.getClass().getMethod("setHeader", String.class, String.class);
			Method getOutputStream = response.getClass().getMethod("getOutputStream");

			setContentType.setAccessible(true);
			setHeader.setAccessible(true);
			getOutputStream.setAccessible(true);

			setContentType.invoke(response, "application/octet-stream; charset=utf-8");
			String encodedFilename = java.net.URLEncoder.encode(filename, "UTF-8").replaceAll("\\+", "%20");
			setHeader.invoke(response, "Content-Disposition", "attachment; filename*=UTF-8''" + encodedFilename);

			out = new BufferedOutputStream((OutputStream) getOutputStream.invoke(response));
			out.write(bos.toByteArray());
			out.flush();
		} catch (Exception e) {
			throw new IOException("Download pdf failed", e);
		} finally {
			if (bos != null) {
				try { bos.close(); } catch (IOException ignore) {}
			}
			if (out != null) {
				try { out.close(); } catch (IOException ignore) {}
			}
		}
	}

	/**
	 * 合并PDF
	 * @param list          需要合并的文件的全路径list
	 * @param destPath      合并后的文件存储的全路径
	 * @throws Exception
	 */
	public static void mergePdf(List<String> list, String destPath) throws Exception {
		new MergeHelper().mergePdf(list, destPath);
	}

	/**
	 * 根据html内容创建Pdf，并输出到指定路径
	 * @param document
	 * @param html         html内容
	 * @param destPath     文件写到哪里的全路径，例如：D:\\1.pdf
	 * @return             返回生成了多少页
	 * @throws Exception
	 */
	public static int createPdf(Document document, String html, String destPath) throws Exception {
		return createPdf(document, html, destPath, null);
	}

	/**
	 * 根据html内容创建Pdf，并输出到指定路径
	 * @param document
	 * @param html         html内容
	 * @param destPath     文件写到哪里的全路径，例如：D:\\1.pdf
	 * @param fontFamily   自定义字体，可以是绝对路径、相对路径、resources下的路径，例如：resources:fonts/simsun.ttc
	 * @return             返回生成了多少页
	 * @throws Exception
	 */
	public static int createPdf(Document document, String html, String destPath, String fontFamily) throws Exception {
		if (document==null) {
			throw new IllegalArgumentException("PDF Document不能为空");
		}
		if (html==null) {
			throw new IllegalArgumentException("HTML内容不能为空");
		}
		if (destPath==null || destPath.trim().length()==0) {
			throw new IllegalArgumentException("PDF输出路径不能为空");
		}

		File targetFile = new File(destPath);
		File parentFile = targetFile.getParentFile();
		if (parentFile!=null && !parentFile.exists()) {
			parentFile.mkdirs();
		}

		try (FileOutputStream out = new FileOutputStream(targetFile)) {
			PdfWriter writer = PdfWriter.getInstance(document, out);
			document.open();

			XMLWorkerHelper worker = XMLWorkerHelper.getInstance();
			worker.parseXHtml(writer, document, new ByteArrayInputStream(html.getBytes(StandardCharsets.UTF_8)), null, new FontFamilyHelper(fontFamily));

			document.close();

			return writer.getPageNumber();
		} finally {
			if (document.isOpen()) {
				document.close();
			}
		}
	}

	/**
	 * 设置水印
	 * @param reader
	 * @param obj       水印内容，可以是纯英文文字，如果是中文的话，必须使用cn.javaex.officejj.common.entity.Font
	 * @param destPath
	 * @throws Exception
	 */
	public static void setWatermark(PdfReader reader, Object obj, String destPath) throws Exception {
		new PdfHelper().setWatermark(reader, obj, destPath);
	}

	/**
	 * 按页拆分PDF。
	 * 每一页会输出成一个独立PDF文件。
	 * @param reader PDF读取器
	 * @param destDir 输出目录
	 * @param filePrefix 文件名前缀
	 * @return 输出文件路径集合
	 * @throws Exception
	 */
	public static List<String> splitPdf(PdfReader reader, String destDir, String filePrefix) throws Exception {
		if (reader==null) {
			throw new IllegalArgumentException("PDF Reader不能为空");
		}
		if (destDir==null || destDir.trim().length()==0) {
			throw new IllegalArgumentException("PDF输出目录不能为空");
		}
		File dir = new File(destDir);
		if (!dir.exists()) {
			dir.mkdirs();
		}
		if (filePrefix==null || filePrefix.length()==0) {
			filePrefix = "page";
		}

		List<String> fileList = new ArrayList<String>();
		try {
			for (int i=1; i<=reader.getNumberOfPages(); i++) {
				String filePath = new File(dir, filePrefix + "_" + i + ".pdf").getAbsolutePath();
				copyPages(reader, new int[] {i}, filePath);
				fileList.add(filePath);
			}
			return fileList;
		} finally {
			reader.close();
		}
	}

	/**
	 * 抽取指定页生成新PDF。
	 * @param reader PDF读取器
	 * @param pages 页码数组，从1开始
	 * @param destPath 输出路径
	 * @throws Exception
	 */
	public static void extractPages(PdfReader reader, int[] pages, String destPath) throws Exception {
		try {
			copyPages(reader, pages, destPath);
		} finally {
			if (reader!=null) {
				reader.close();
			}
		}
	}

	/**
	 * 删除指定页生成新PDF。
	 * @param reader PDF读取器
	 * @param deletePages 需要删除的页码，从1开始
	 * @param destPath 输出路径
	 * @throws Exception
	 */
	public static void deletePages(PdfReader reader, int[] deletePages, String destPath) throws Exception {
		if (reader==null) {
			throw new IllegalArgumentException("PDF Reader不能为空");
		}
		List<Integer> keepPages = new ArrayList<Integer>();
		for (int i=1; i<=reader.getNumberOfPages(); i++) {
			if (!contains(deletePages, i)) {
				keepPages.add(i);
			}
		}
		int[] pages = new int[keepPages.size()];
		for (int i=0; i<keepPages.size(); i++) {
			pages[i] = keepPages.get(i);
		}
		extractPages(reader, pages, destPath);
	}

	/**
	 * 按指定页码顺序重排PDF。
	 * @param reader PDF读取器
	 * @param pages 新顺序页码，从1开始
	 * @param destPath 输出路径
	 * @throws Exception
	 */
	public static void reorderPages(PdfReader reader, int[] pages, String destPath) throws Exception {
		extractPages(reader, pages, destPath);
	}

	/**
	 * 插入另一个PDF。
	 * @param reader 原PDF读取器
	 * @param insertReader 待插入PDF读取器
	 * @param afterPage 插入到第几页之后；0表示插入到开头
	 * @param destPath 输出路径
	 * @throws Exception
	 */
	public static void insertPdf(PdfReader reader, PdfReader insertReader, int afterPage, String destPath) throws Exception {
		if (reader==null || insertReader==null) {
			throw new IllegalArgumentException("PDF Reader不能为空");
		}
		ensureParent(destPath);
		Document document = new Document();
		PdfCopy copy = null;
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(destPath);
			copy = new PdfCopy(document, out);
			document.open();
			for (int i=1; i<=reader.getNumberOfPages(); i++) {
				if (afterPage==0 && i==1) {
					copyAllPages(insertReader, copy);
				}
				copy.addPage(copy.getImportedPage(reader, i));
				if (i==afterPage) {
					copyAllPages(insertReader, copy);
				}
			}
			if (afterPage>=reader.getNumberOfPages()) {
				copyAllPages(insertReader, copy);
			}
		} finally {
			if (document.isOpen()) {
				document.close();
			}
			if (out!=null) {
				try { out.close(); } catch (Exception ignore) {}
			}
			reader.close();
			insertReader.close();
		}
	}

	/**
	 * 在PDF指定页写入文字。
	 * @param reader PDF读取器
	 * @param destPath 输出路径
	 * @param text 文本
	 * @param pageNum 页码，从1开始；0表示所有页
	 * @param x x坐标
	 * @param y y坐标
	 * @throws Exception
	 */
	public static void stampText(PdfReader reader, String destPath, String text, int pageNum, float x, float y) throws Exception {
		stampText(reader, destPath, text, pageNum, x, y, 12F);
	}

	/**
	 * 在PDF指定页写入文字。
	 * @param reader PDF读取器
	 * @param destPath 输出路径
	 * @param text 文本
	 * @param pageNum 页码，从1开始；0表示所有页
	 * @param x x坐标
	 * @param y y坐标
	 * @param fontSize 字号
	 * @throws Exception
	 */
	public static void stampText(PdfReader reader, String destPath, String text, int pageNum, float x, float y, float fontSize) throws Exception {
		ensureParent(destPath);
		PdfStamper stamper = null;
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(destPath);
			stamper = new PdfStamper(reader, out);
			BaseFont baseFont = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
			int start = pageNum<=0 ? 1 : pageNum;
			int end = pageNum<=0 ? reader.getNumberOfPages() : pageNum;
			for (int i=start; i<=end; i++) {
				PdfContentByte content = stamper.getOverContent(i);
				content.beginText();
				content.setFontAndSize(baseFont, fontSize);
				content.showTextAligned(Element.ALIGN_LEFT, text==null ? "" : text, x, y, 0);
				content.endText();
			}
		} finally {
			if (stamper!=null) {
				try { stamper.close(); } catch (Exception ignore) {}
			}
			if (out!=null) {
				try { out.close(); } catch (Exception ignore) {}
			}
			if (reader!=null) {
				reader.close();
			}
		}
	}

	/**
	 * 在PDF指定页写入图片。
	 * @param reader PDF读取器
	 * @param destPath 输出路径
	 * @param imagePath 图片路径
	 * @param pageNum 页码，从1开始；0表示所有页
	 * @param x x坐标
	 * @param y y坐标
	 * @param width 图片宽度
	 * @param height 图片高度
	 * @throws Exception
	 */
	public static void stampImage(PdfReader reader, String destPath, String imagePath, int pageNum, float x, float y, float width, float height) throws Exception {
		ensureParent(destPath);
		PdfStamper stamper = null;
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(destPath);
			stamper = new PdfStamper(reader, out);
			Image image = Image.getInstance(imagePath);
			image.scaleAbsolute(width, height);
			image.setAbsolutePosition(x, y);
			int start = pageNum<=0 ? 1 : pageNum;
			int end = pageNum<=0 ? reader.getNumberOfPages() : pageNum;
			for (int i=start; i<=end; i++) {
				stamper.getOverContent(i).addImage(image);
			}
		} finally {
			if (stamper!=null) {
				try { stamper.close(); } catch (Exception ignore) {}
			}
			if (out!=null) {
				try { out.close(); } catch (Exception ignore) {}
			}
			if (reader!=null) {
				reader.close();
			}
		}
	}

	/**
	 * 添加页码。
	 * @param reader PDF读取器
	 * @param destPath 输出路径
	 * @param format 页码格式，使用 {page} 和 {total} 占位，例如：第 {page} / {total} 页
	 * @throws Exception
	 */
	public static void addPageNumber(PdfReader reader, String destPath, String format) throws Exception {
		ensureParent(destPath);
		PdfStamper stamper = null;
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(destPath);
			stamper = new PdfStamper(reader, out);
			BaseFont baseFont = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
			int total = reader.getNumberOfPages();
			for (int i=1; i<=total; i++) {
				Rectangle pageSize = reader.getPageSize(i);
				String text = (format==null || format.length()==0 ? "{page}/{total}" : format).replace("{page}", String.valueOf(i)).replace("{total}", String.valueOf(total));
				PdfContentByte content = stamper.getOverContent(i);
				content.beginText();
				content.setFontAndSize(baseFont, 10F);
				content.showTextAligned(Element.ALIGN_CENTER, text, pageSize.getWidth() / 2, 24F, 0);
				content.endText();
			}
		} finally {
			if (stamper!=null) {
				try { stamper.close(); } catch (Exception ignore) {}
			}
			if (out!=null) {
				try { out.close(); } catch (Exception ignore) {}
			}
			if (reader!=null) {
				reader.close();
			}
		}
	}

	/**
	 * 添加页眉页脚。
	 * @param reader PDF读取器
	 * @param destPath 输出路径
	 * @param headerText 页眉文本，允许为空
	 * @param footerText 页脚文本，允许为空
	 * @throws Exception
	 */
	public static void addHeaderFooter(PdfReader reader, String destPath, String headerText, String footerText) throws Exception {
		ensureParent(destPath);
		PdfStamper stamper = null;
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(destPath);
			stamper = new PdfStamper(reader, out);
			BaseFont baseFont = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
			for (int i=1; i<=reader.getNumberOfPages(); i++) {
				Rectangle pageSize = reader.getPageSize(i);
				PdfContentByte content = stamper.getOverContent(i);
				content.beginText();
				content.setFontAndSize(baseFont, 10F);
				if (headerText!=null && headerText.length()>0) {
					content.showTextAligned(Element.ALIGN_CENTER, headerText, pageSize.getWidth() / 2, pageSize.getHeight() - 24F, 0);
				}
				if (footerText!=null && footerText.length()>0) {
					content.showTextAligned(Element.ALIGN_CENTER, footerText, pageSize.getWidth() / 2, 24F, 0);
				}
				content.endText();
			}
		} finally {
			if (stamper!=null) {
				try { stamper.close(); } catch (Exception ignore) {}
			}
			if (out!=null) {
				try { out.close(); } catch (Exception ignore) {}
			}
			if (reader!=null) {
				reader.close();
			}
		}
	}

	/**
	 * 加密PDF。
	 * @param reader PDF读取器
	 * @param destPath 输出路径
	 * @param userPassword 用户密码
	 * @param ownerPassword 所有者密码
	 * @throws Exception
	 */
	public static void encryptPdf(PdfReader reader, String destPath, String userPassword, String ownerPassword) throws Exception {
		encryptPdf(reader, destPath, userPassword, ownerPassword, PdfWriter.ALLOW_PRINTING);
	}

	/**
	 * 加密PDF并设置权限。
	 * permissions可使用PdfWriter.ALLOW_PRINTING、ALLOW_COPY、ALLOW_MODIFY_CONTENTS等常量组合。
	 * @param reader PDF读取器
	 * @param destPath 输出路径
	 * @param userPassword 用户密码
	 * @param ownerPassword 所有者密码
	 * @param permissions 权限位
	 * @throws Exception
	 */
	public static void encryptPdf(PdfReader reader, String destPath, String userPassword, String ownerPassword, int permissions) throws Exception {
		ensureParent(destPath);
		PdfStamper stamper = null;
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(destPath);
			stamper = new PdfStamper(reader, out);
			stamper.setEncryption(
					userPassword==null ? null : userPassword.getBytes(StandardCharsets.UTF_8),
					ownerPassword==null ? null : ownerPassword.getBytes(StandardCharsets.UTF_8),
					permissions,
					PdfWriter.ENCRYPTION_AES_128);
		} finally {
			if (stamper!=null) {
				try { stamper.close(); } catch (Exception ignore) {}
			}
			if (out!=null) {
				try { out.close(); } catch (Exception ignore) {}
			}
			if (reader!=null) {
				reader.close();
			}
		}
	}

	/**
	 * 读取PDF表单字段和值。
	 * @param reader PDF读取器
	 * @return
	 */
	public static Map<String, String> getFormFields(PdfReader reader) {
		if (reader==null) {
			throw new IllegalArgumentException("PDF Reader不能为空");
		}
		try {
			AcroFields fields = reader.getAcroFields();
			Map<String, String> map = new LinkedHashMap<String, String>();
			for (String name : fields.getFields().keySet()) {
				map.put(name, fields.getField(name));
			}
			return map;
		} finally {
			reader.close();
		}
	}

	/**
	 * 扁平化PDF表单。
	 * 扁平化后表单值会成为普通页面内容，用户不能继续编辑表单控件。
	 * @param reader PDF读取器
	 * @param destPath 输出路径
	 * @throws Exception
	 */
	public static void flattenPdf(PdfReader reader, String destPath) throws Exception {
		ensureParent(destPath);
		PdfStamper stamper = null;
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(destPath);
			stamper = new PdfStamper(reader, out);
			stamper.setFormFlattening(true);
		} finally {
			if (stamper!=null) {
				try { stamper.close(); } catch (Exception ignore) {}
			}
			if (out!=null) {
				try { out.close(); } catch (Exception ignore) {}
			}
			if (reader!=null) {
				reader.close();
			}
		}
	}

	/**
	 * 对PDF进行数字签名。
	 * 该方法只负责PDF签名写入，证书链和私钥由业务系统从证书文件、证书服务或硬件Key中获取。
	 * @param reader PDF读取器
	 * @param destPath 输出路径
	 * @param privateKey 私钥
	 * @param chain 证书链
	 * @param reason 签名原因，允许为空
	 * @param location 签名地点，允许为空
	 * @throws Exception
	 */
	public static void signPdf(PdfReader reader, String destPath, PrivateKey privateKey, Certificate[] chain, String reason, String location) throws Exception {
		signPdf(reader, destPath, privateKey, chain, reason, location, null, 1);
	}

	/**
	 * 对PDF进行可见数字签名。
	 * @param reader PDF读取器
	 * @param destPath 输出路径
	 * @param privateKey 私钥
	 * @param chain 证书链
	 * @param reason 签名原因，允许为空
	 * @param location 签名地点，允许为空
	 * @param visibleRect 签章显示区域，传null表示不可见签名
	 * @param pageNum 可见签章页码，从1开始
	 * @throws Exception
	 */
	public static void signPdf(PdfReader reader, String destPath, PrivateKey privateKey, Certificate[] chain,
			String reason, String location, Rectangle visibleRect, int pageNum) throws Exception {
		if (reader==null) {
			throw new IllegalArgumentException("PDF Reader不能为空");
		}
		if (privateKey==null) {
			throw new IllegalArgumentException("签名私钥不能为空");
		}
		if (chain==null || chain.length==0) {
			throw new IllegalArgumentException("签名证书链不能为空");
		}
		ensureParent(destPath);
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(destPath);
			PdfStamper stamper = PdfStamper.createSignature(reader, out, '\0');
			PdfSignatureAppearance appearance = stamper.getSignatureAppearance();
			appearance.setReason(reason);
			appearance.setLocation(location);
			if (visibleRect!=null) {
				appearance.setVisibleSignature(visibleRect, pageNum, "officejj_signature");
			}

			ExternalDigest digest = new BouncyCastleDigest();
			ExternalSignature signature = new PrivateKeySignature(privateKey, DigestAlgorithms.SHA256, null);
			MakeSignature.signDetached(appearance, digest, signature, chain, null, null, null, 0, MakeSignature.CryptoStandard.CADES);
		} finally {
			if (out!=null) {
				try { out.close(); } catch (Exception ignore) {}
			}
			if (reader!=null) {
				reader.close();
			}
		}
	}

	/**
	 * 复制指定页。
	 * @param reader PDF读取器
	 * @param pages 页码数组
	 * @param destPath 输出路径
	 * @throws Exception
	 */
	private static void copyPages(PdfReader reader, int[] pages, String destPath) throws Exception {
		if (reader==null) {
			throw new IllegalArgumentException("PDF Reader不能为空");
		}
		if (pages==null || pages.length==0) {
			throw new IllegalArgumentException("页码不能为空");
		}
		ensureParent(destPath);
		Document document = new Document();
		PdfCopy copy = null;
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(destPath);
			copy = new PdfCopy(document, out);
			document.open();
			for (int page : pages) {
				if (page<=0 || page>reader.getNumberOfPages()) {
					throw new IllegalArgumentException("页码超出范围：" + page);
				}
				PdfImportedPage importedPage = copy.getImportedPage(reader, page);
				copy.addPage(importedPage);
			}
		} finally {
			if (document.isOpen()) {
				document.close();
			}
			if (out!=null) {
				try { out.close(); } catch (Exception ignore) {}
			}
		}
	}

	private static void copyAllPages(PdfReader reader, PdfCopy copy) throws Exception {
		for (int i=1; i<=reader.getNumberOfPages(); i++) {
			copy.addPage(copy.getImportedPage(reader, i));
		}
	}

	private static boolean contains(int[] arr, int value) {
		if (arr==null) {
			return false;
		}
		for (int item : arr) {
			if (item==value) {
				return true;
			}
		}
		return false;
	}

	private static void ensureParent(String destPath) {
		if (destPath==null || destPath.trim().length()==0) {
			throw new IllegalArgumentException("PDF输出路径不能为空");
		}
		File targetFile = new File(destPath);
		File parentFile = targetFile.getParentFile();
		if (parentFile!=null && !parentFile.exists()) {
			parentFile.mkdirs();
		}
	}
}
