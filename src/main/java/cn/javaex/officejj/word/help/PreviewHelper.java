package cn.javaex.officejj.word.help;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

import fr.opensagres.poi.xwpf.converter.core.ImageManager;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;

/**
 * 预览
 * 
 * @author 陈霓清
 */
public class PreviewHelper {

	/**
	 * word转html
	 * @param filePath     word文件绝对路径
	 * @return
	 * @throws Exception
	 */
	public String wordToHtml(String filePath) throws Exception {
		if (filePath==null || filePath.length()==0) {
			return "";
		}
		
		String htmlPath = "";
		
		filePath = filePath.toLowerCase();
		
		if (filePath.endsWith(".docx")) {
			htmlPath = this.docxToHtml(filePath);
		}
		else if (filePath.endsWith(".doc")) {
			htmlPath = this.docToHtml(filePath);
		}
		
		return htmlPath;
	}
	
	/**
	 * 将word（.doc后缀）转为html
	 * @param wordPath word文件的全路径，例如："D:\\Temp\\1.doc"
	 * @return 返回生成的html文件的全路径，例如："D:\\Temp\\1_html\\1.html"
	 * @throws Exception
	 */
	private String docToHtml(String filePath) throws Exception {
		String htmlPath = "";
		
		InputStream in = null;
		OutputStream out = null;
		HWPFDocument doc = null;
		
		try {
			File wordFile = new File(filePath);
			if (!wordFile.exists()) {
				throw new FileNotFoundException("指定文件不存在：" + filePath);
			}
			
			String wordName = wordFile.getName();
			String htmlName = wordName.replace(".doc", ".html");
			String wordFolderPath = wordFile.getParent();
			String htmlFolderPath = wordFolderPath + File.separator + wordName.replace(".doc", "") + "_html";
			
			// 判断html文件是否已存在
			File htmlFile = new File(htmlFolderPath + File.separator + htmlName);
			if (htmlFile.exists()) {
				return htmlFile.getAbsolutePath();
			} else {
				// 生成html文件上级文件夹
				File folder = new File(htmlFolderPath);
				if (!folder.exists()) {
					folder.mkdirs();
				}
			}
			
			// 图片目录
			final String IMAGE_FOLDER_PATH = htmlFolderPath + File.separator + "image";
			
			// 原word文档
			in = new FileInputStream(wordFile);
			doc = new HWPFDocument(in);
			WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
			// 设置图片存放的位置
			wordToHtmlConverter.setPicturesManager(new PicturesManager() {
				public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches, float heightInches) {
					File imgPath = new File(IMAGE_FOLDER_PATH);
					if (!imgPath.exists()) {
						imgPath.mkdirs();
					}
					File file = new File(IMAGE_FOLDER_PATH + File.separator + suggestedName);
					try {
						OutputStream os = new FileOutputStream(file);
						os.write(content);
						os.close();
					} catch (FileNotFoundException e) {
						e.printStackTrace();
					} catch (IOException e) {
						e.printStackTrace();
					}
					// 图片在html文件上的相对路径
					return "image" + File.separator + suggestedName;
				}
			});
			
			// 解析word文档
			wordToHtmlConverter.processDocument(doc);
			Document htmlDocument = wordToHtmlConverter.getDocument();
			out = new FileOutputStream(htmlFile);
			DOMSource domSource = new DOMSource(htmlDocument);
			StreamResult streamResult = new StreamResult(out);
			TransformerFactory factory = TransformerFactory.newInstance();
			Transformer serializer = factory.newTransformer();
			serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
			serializer.setOutputProperty(OutputKeys.INDENT, "yes");    // 是否添加空格
			serializer.setOutputProperty(OutputKeys.METHOD, "html");
			serializer.transform(domSource, streamResult);
			
			htmlPath = htmlFile.getAbsolutePath();
		} catch (FileNotFoundException e) {
			throw new FileNotFoundException("指定文件不存在：" + filePath);
		} catch (IOException e) {
			throw new IOException(e);
		} catch (Exception e) {
			throw new Exception(e);
		} finally {
			IOUtils.closeQuietly(doc);
			IOUtils.closeQuietly(out);
			IOUtils.closeQuietly(in);
		}
		
		return htmlPath;
	}

	/**
	 * 将word（.docx后缀）转为html
	 * @param wordPath word文件的全路径，例如："D:\\Temp\\1.docx"
	 * @return 返回生成的html文件的全路径，例如："D:\\Temp\\1_html\\1.html"
	 * @throws Exception
	 */
	public String docxToHtml(String filePath) throws Exception {
		String htmlPath = "";
		
		InputStream in = null;
		OutputStream out = null;
		XWPFDocument word = null;
		
		try {
			File wordFile = new File(filePath);
			if (!wordFile.exists()) {
				throw new FileNotFoundException("指定文件不存在：" + filePath);
			}
			
			String wordName = wordFile.getName();
			String htmlName = wordName.replace(".docx", ".html");
			String wordFolderPath = wordFile.getParent();
			String htmlFolderPath = wordFolderPath + File.separator + wordName.replace(".docx", "") + "_html";
			
			// 1.0 判断html文件是否已存在
			File htmlFile = new File(htmlFolderPath + File.separator + htmlName);
			if (htmlFile.exists()) {
				return htmlFile.getAbsolutePath();
			} else {
				// 生成html文件上级文件夹
				File folder = new File(htmlFolderPath);
				if (!folder.exists()) {
					folder.mkdirs();
				}
			}
			
			// 2.0 生成html文件
			// 2.1 读取word
			word = new XWPFDocument(new FileInputStream(wordFile));
			// 2.2 解析 XHTML配置
			ImageManager imageManager = new ImageManager(new File(htmlFolderPath), "image");    // html中图片的路径 相对路径
			
			XHTMLOptions options = XHTMLOptions.create();
			options.setImageManager(imageManager);
			options.setIgnoreStylesIfUnused(false);
			options.setFragment(true);
			
			// 2.3 将 XWPFDocument转换成XHTML
			out = new FileOutputStream(htmlFile);
			XHTMLConverter.getInstance().convert(word, out, options);
			
			htmlPath = htmlFile.getAbsolutePath();
		} catch (FileNotFoundException e) {
			throw new FileNotFoundException("指定文件不存在：" + filePath);
		} catch (IOException e) {
			throw new IOException(e);
		} catch (Exception e) {
			throw new Exception(e);
		} finally {
			IOUtils.closeQuietly(word);
			IOUtils.closeQuietly(out);
			IOUtils.closeQuietly(in);
		}
		
		return htmlPath;
	}

}
