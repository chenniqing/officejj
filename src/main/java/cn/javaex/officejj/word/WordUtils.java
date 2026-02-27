package cn.javaex.officejj.word;

import java.io.BufferedOutputStream;
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
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import cn.javaex.officejj.common.util.PathHandler;
import cn.javaex.officejj.word.help.MergeHelper;
import cn.javaex.officejj.word.help.ParagraphHelper;
import cn.javaex.officejj.word.help.TableHelper;
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
		return getDocx(new FileInputStream(filePath));
	}
	
	/**
	 * 读取resources文件夹下的Word
	 * @param filePath      resources文件夹下的路径，例如：template/word/模板.docx
	 * @return
	 * @throws IOException 
	 */
	public static XWPFDocument getDocxFromResource(String filePath) throws IOException {
		InputStream in = PathHandler.getInputStreamFromResource(filePath);
		return getDocx(in);
	}
	
	/**
	 * 通过流读取Word
	 * @param in
	 * @return
	 */
	public static XWPFDocument getDocx(InputStream in) {
		XWPFDocument word = null;
		
		try {
			word = new XWPFDocument(in);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(in);
		}
		
		return word;
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
	 * 输出Word到指定路径
	 * @param word         word对象，支持 doc 和 docx，例如：XWPFDocument word
	 * @param filePath     文件写到哪里的全路径，例如：D:\\1.docx
	 */
	public static void output(POIXMLDocument word, String filePath) {
		// 保证这个文件的父文件夹必须要存在
		File targetFile = new File(filePath);
		if (!targetFile.getParentFile().exists()) {
			targetFile.getParentFile().mkdirs();
		}
		
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(targetFile);
			word.write(out);
			out.flush();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (out != null) {
				try { out.close(); } catch (IOException ignore) {}
			}
			try { word.close(); } catch (Exception ignore) {}
		}
	}
	
	/**
	 * 下载Word（兼容 javax 和 jakarta Servlet 环境）
	 * @param response    javax.servlet.http.HttpServletResponse 或 jakarta.servlet.http.HttpServletResponse
	 * @param word        POIXMLDocument（如 XWPFDocument 或 HWPFDocument 等）
	 * @param filename    带后缀的文件名，例如："test.docx"
	 */
	public static void download(Object response, POIXMLDocument word, String filename) throws IOException {
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
			e.printStackTrace();
			throw new IOException("Download word failed", e);
		} finally {
			if (out != null) {
				try { out.close(); } catch (IOException ignore) {}
			}
			try { word.close(); } catch (Exception ignore) {}
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
	
}
