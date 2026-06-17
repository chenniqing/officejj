package cn.javaex.officejj.word;

import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import cn.javaex.officejj.common.util.PathHandler;
import cn.javaex.officejj.word.help.MergeHelper;
import cn.javaex.officejj.word.help.ParagraphHelper;
import cn.javaex.officejj.word.help.TableHelper;
import cn.javaex.officejj.word.help.WordTemplateHelper;
import cn.javaex.officejj.word.help.WordHelper;

/**
 * Word工具类
 *
 * @author 陈霓清
 */
public class WordUtils {

	/**
	 * 通过路径读取Word
	 * @param filePath     例如：D:\\123.docx
	 * @return
	 * @throws FileNotFoundException
	 */
	public static XWPFDocument getDocx(String filePath) throws FileNotFoundException {
		return getDocx(new FileInputStream(filePath), true);
	}

	/**
	 * 读取resources文件夹下的Word
	 * @param filePath      resources文件夹下的路径，例如：template/word/模板.docx
	 * @return
	 * @throws IOException
	 */
	public static XWPFDocument getDocxFromResource(String filePath) throws IOException {
		InputStream in = PathHandler.getInputStreamFromResource(filePath);
		return getDocx(in, true);
	}

	/**
	 * 通过流读取Word，不关闭调用方传入的输入流。
	 * @param in
	 * @return
	 */
	public static XWPFDocument getDocx(InputStream in) {
		return getDocx(in, false);
	}

	/**
	 * 通过流读取Word。
	 * 工具类内部创建的输入流需要关闭；调用方传入的输入流默认由调用方自己关闭。
	 * @param in
	 * @param closeInputStream 是否在读取完成后关闭输入流
	 * @return
	 */
	private static XWPFDocument getDocx(InputStream in, boolean closeInputStream) {
		if (in==null) {
			throw new IllegalArgumentException("Word输入流不能为空");
		}

		try {
			return new XWPFDocument(in);
		} catch (Exception e) {
			throw new RuntimeException("读取Word失败", e);
		} finally {
			if (closeInputStream) {
				try { in.close(); } catch (IOException ignore) {}
			}
		}
	}

	/**
	 * 替换Word中占位符的内容
	 * @param word Word文件对象
	 * @param param 允许为空
	 * @return
	 * @throws Exception
	 */
	public static XWPFDocument writeDocx(XWPFDocument word, Map<String, Object> param) throws Exception {
		if (param!=null && param.size()>0) {
			ParagraphHelper paragraphHelper = new ParagraphHelper();
			TableHelper tableHelper = new TableHelper();

			// 替换段落
			paragraphHelper.replaceParagraph(word, param);
			// 替换表格
			tableHelper.replaceTable(word, param);
		}

		return word;
	}

	/**
	 * 增强模板写入。
	 * 支持嵌套属性、表格行循环和简单条件段落，适合合同、报告、审批单等复杂模板。
	 * @param word Word文件对象
	 * @param param 模板参数
	 * @return
	 * @throws Exception
	 */
	public static XWPFDocument writeDocxTemplate(XWPFDocument word, Map<String, Object> param) throws Exception {
		return new WordTemplateHelper().render(word, param);
	}

	/**
	 * 输出Word到指定路径，默认写完后关闭文档对象。
	 * @param word         word对象，支持 doc 和 docx，例如：XWPFDocument word
	 * @param filePath     文件写到哪里的全路径，例如：D:\\1.docx
	 */
	public static void output(POIXMLDocument word, String filePath) {
		output(word, filePath, true);
	}

	/**
	 * 输出Word到指定路径。
	 * 默认方法会关闭文档对象；需要继续复用文档对象时，可调用本重载并传false。
	 * @param word         word对象，支持 doc 和 docx，例如：XWPFDocument word
	 * @param filePath     文件写到哪里的全路径，例如：D:\\1.docx
	 * @param closeWord    输出后是否关闭文档对象
	 */
	public static void output(POIXMLDocument word, String filePath, boolean closeWord) {
		// 保证这个文件的父文件夹必须要存在
		File targetFile = new File(filePath);
		File parentFile = targetFile.getParentFile();
		if (parentFile!=null && !parentFile.exists()) {
			parentFile.mkdirs();
		}

		FileOutputStream out = null;
		try {
			out = new FileOutputStream(targetFile);
			word.write(out);
			out.flush();
		} catch (Exception e) {
			throw new RuntimeException("输出Word失败：" + filePath, e);
		} finally {
			if (out != null) {
				try { out.close(); } catch (IOException ignore) {}
			}
			if (closeWord && word!=null) {
				try { word.close(); } catch (Exception ignore) {}
			}
		}
	}

	/**
	 * 下载Word（兼容 javax 和 jakarta Servlet 环境），默认写完后关闭文档对象。
	 * @param response    javax.servlet.http.HttpServletResponse 或 jakarta.servlet.http.HttpServletResponse
	 * @param word        POIXMLDocument（如 XWPFDocument 或 HWPFDocument 等）
	 * @param filename    带后缀的文件名，例如："test.docx"
	 */
	public static void download(Object response, POIXMLDocument word, String filename) throws IOException {
		download(response, word, filename, true);
	}

	/**
	 * 下载Word（兼容 javax 和 jakarta Servlet 环境）。
	 * 默认方法会关闭文档对象；需要继续复用文档对象时，可调用本重载并传false。
	 * @param response    javax.servlet.http.HttpServletResponse 或 jakarta.servlet.http.HttpServletResponse
	 * @param word        POIXMLDocument（如 XWPFDocument 或 HWPFDocument 等）
	 * @param filename    带后缀的文件名，例如："test.docx"
	 * @param closeWord   下载后是否关闭文档对象
	 */
	public static void download(Object response, POIXMLDocument word, String filename, boolean closeWord) throws IOException {
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
			word.write(out);
			out.flush();
		} catch (Exception e) {
			throw new IOException("Download word failed", e);
		} finally {
			if (out != null) {
				try { out.close(); } catch (IOException ignore) {}
			}
			if (closeWord && word!=null) {
				try { word.close(); } catch (Exception ignore) {}
			}
		}
	}

	/**
	 * 将Word写入字节数组，不关闭传入的文档对象，便于调用方自行决定生命周期。
	 * @param word
	 * @return
	 * @throws IOException
	 */
	public static byte[] toByteArray(POIXMLDocument word) throws IOException {
		try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
			word.write(bos);
			return bos.toByteArray();
		}
	}

	/**
	 * 合并Word（分页）
	 * @param word1
	 * @param word2
	 * @return
	 * @throws Exception
	 */
	public static XWPFDocument mergeDocx(XWPFDocument word1, XWPFDocument word2) throws Exception {
		return new MergeHelper().mergeDocx(word1, word2);
	}

	/**
	 * 合并Word
	 * @param word1
	 * @param word2
	 * @param isPage    是否分页
	 * @return
	 * @throws Exception
	 */
	public static XWPFDocument mergeDocx(XWPFDocument word1, XWPFDocument word2, boolean isPage) throws Exception {
		return new MergeHelper().mergeDocx(word1, word2, isPage);
	}

	/**
	 * 合并Word（分页）
	 * @param list         word绝对路径集合
	 * @param destPath     输出路径，例如：D:\\Temp\\合并.docx
	 * @throws Exception
	 */
	public static void mergeDocx(List<String> list, String destPath) throws Exception {
		new MergeHelper().mergeDocx(list, destPath);
	}

	/**
	 * 合并Word
	 * @param list         word绝对路径集合
	 * @param destPath     输出路径，例如：D:\\Temp\\合并.docx
	 * @param isPage       是否分页
	 * @throws Exception
	 */
	public static void mergeDocx(List<String> list, String destPath, boolean isPage) throws Exception {
		new MergeHelper().mergeDocx(list, destPath, isPage);
	}

	/**
	 * 设置水印
	 * @param word
	 * @param content     水印文字内容
	 */
	public static void setWatermark(XWPFDocument word, String content) {
		new WordHelper().setWatermark(word, content);
	}

	/**
	 * 设置只读
	 * @param word
	 */
	public static void setReadOnly(XWPFDocument word) {
		word.enforceReadonlyProtection(UUID.randomUUID().toString().replace("-", ""), HashAlgorithm.sha512);
	}

	/**
	 * 设置只读
	 * @param word
	 * @param password    密码
	 */
	public static void setReadOnly(XWPFDocument word, String password) {
		word.enforceReadonlyProtection(password, HashAlgorithm.sha512);
	}

	/**
	 * 设置页眉
	 * @param word
	 * @param obj    字符串
	 *               cn.javaex.officejj.common.entity.Font
	 *               cn.javaex.officejj.common.entity.Picture
	 * @param align
	 */
	public static void setHeader(XWPFDocument word, Object obj, ParagraphAlignment align) {
		new WordHelper().setHeader(word, obj, align);
	}

	/**
	 * 设置页眉（左右两端对齐）
	 * @param word
	 * @param obj1    字符串
	 *                cn.javaex.officejj.common.entity.Font
	 *                cn.javaex.officejj.common.entity.Picture
	 * @param obj2    字符串
	 *                cn.javaex.officejj.common.entity.Font
	 *                cn.javaex.officejj.common.entity.Picture
	 * @param spacing  间距，默认word下请填写 1440
	 */
	public static void setHeader(XWPFDocument word, Object obj1, Object obj2, int spacing) {
		new WordHelper().setHeader(word, obj1, obj2, spacing);
	}

	/**
	 * 设置页脚
	 * @param word
	 * @param obj    字符串
	 *               cn.javaex.officejj.common.entity.Font
	 *               cn.javaex.officejj.common.entity.Picture
	 * @param align
	 */
	public static void setFooter(XWPFDocument word, Object obj, ParagraphAlignment align) {
		new WordHelper().setFooter(word, obj, align);
	}

	/**
	 * 设置页脚（左右两端对齐）
	 * @param word
	 * @param obj1    字符串
	 *                cn.javaex.officejj.common.entity.Font
	 *                cn.javaex.officejj.common.entity.Picture
	 * @param obj2    字符串
	 *                cn.javaex.officejj.common.entity.Font
	 *                cn.javaex.officejj.common.entity.Picture
	 * @param spacing  间距，默认word下请填写 1440
	 */
	public static void setFooter(XWPFDocument word, Object obj1, Object obj2, int spacing) {
		new WordHelper().setFooter(word, obj1, obj2, spacing);
	}

	/**
	 * 设置页码
	 * @param word
	 * @param align
	 */
	public static void setPageNumber(XWPFDocument word, ParagraphAlignment align) {
		new WordHelper().setPageNumber(word, align);
	}

	/**
	 * 设置书签文本。
	 * @param word Word文档
	 * @param bookmarkName 书签名称
	 * @param text 写入文本
	 */
	public static void setBookmarkText(XWPFDocument word, String bookmarkName, String text) {
		new WordTemplateHelper().setBookmarkText(word, bookmarkName, text);
	}

	/**
	 * 给段落追加超链接。
	 * @param paragraph 段落
	 * @param text 显示文本
	 * @param url 链接地址
	 * @throws Exception
	 */
	public static void addHyperlink(XWPFParagraph paragraph, String text, String url) throws Exception {
		new WordTemplateHelper().addHyperlink(paragraph, text, url);
	}

	/**
	 * 插入简单目录域。
	 * Word打开文档后可更新域得到目录内容。
	 * @param paragraph 段落
	 */
	public static void addTocField(XWPFParagraph paragraph) {
		new WordTemplateHelper().addTocField(paragraph);
	}

}
