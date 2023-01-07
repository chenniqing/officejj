package cn.javaex.officejj.pdf;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.util.IOUtils;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;

import cn.javaex.officejj.common.util.FileHandler;
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
		return getPdf(in);
	}
	
	/**
	 * 通过流读取Pdf
	 * @param in
	 * @return
	 */
	public static PdfReader getPdf(InputStream in) {
		PdfReader reader = null;
		
		try {
			reader = new PdfReader(in);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(in);
		}
		
		return reader;
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
		
		FileOutputStream out = null;
		ByteArrayOutputStream bos = null;
		
		try {
			bos = new ByteArrayOutputStream();
			PdfStamper stamper = new PdfStamper(reader, bos);
			AcroFields form = stamper.getAcroFields();
			
			// 替换占位符为数据中的内容
			form = acroFieldsHelp.replaceContent(form, stamper, param);
			
			stamper.setFormFlattening(readOnly);    // 如果为false那么生成的PDF文件还能编辑
			stamper.close();
			
			reader.close();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(out);
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
		if (!targetFile.getParentFile().exists()) {
			targetFile.getParentFile().mkdirs();
		}
		OutputStream out = null;
		try {
			out = new FileOutputStream(targetFile);
			out.write(bos.toByteArray());
			out.flush();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(bos);
			IOUtils.closeQuietly(out);
		}
	}
	
	/**
	 * 下载Pdf
	 * @param fileName     带后缀的文件名，例如："test.pdf"
	 * @throws IOException
	 */
	public static void download(HttpServletResponse response, ByteArrayOutputStream bos, String fileName) throws IOException {
		String folderPath = PathHandler.getFolderPath();
		
		String fileUrl = folderPath + File.separator + fileName;
		
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(fileUrl);
			out.write(bos.toByteArray()); 
			out.flush();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(bos);
			IOUtils.closeQuietly(out);
		}
		
		FileHandler.downloadFile(response, fileUrl);
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
	 * @param fontFamily   自定义字体，可以是绝对路径、相对路径、resources下的路径，例如：resources:fonts/simsun.ttc
	 * @return             返回生成了多少页
	 * @throws Exception 
	 */
	public static int createPdf(Document document, String html, String destPath, String fontFamily) throws Exception {
		PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(destPath));
		
		document.open();
		
		XMLWorkerHelper worker = XMLWorkerHelper.getInstance();
		worker.parseXHtml(writer, document, new ByteArrayInputStream(html.getBytes()), null, new FontFamilyHelper(fontFamily));
		
		document.close();
		
		return writer.getPageNumber();
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
}
