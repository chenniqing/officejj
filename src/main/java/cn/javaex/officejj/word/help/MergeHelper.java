package cn.javaex.officejj.word.help;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import cn.javaex.officejj.word.WordUtils;

/**
 * 合并Word
 * 
 * @author 陈霓清
 */
public class MergeHelper {

	/**
	 * 合并word(.docx后缀)
	 * @param list
	 * @param destPath
	 * @throws Exception 
	 */
	public void mergeDocx(List<String> list, String destPath) throws Exception {
		this.mergeDocx(list, destPath, true);
	}
	
	/**
	 * 合并word(.docx后缀)
	 * @param list
	 * @param destPath
	 * @param isPage           是否分页
	 * @throws Exception 
	 */
	public void mergeDocx(List<String> list, String destPath, boolean isPage) throws Exception {
		if (list==null || list.isEmpty()) {
			return;
		}
		
		if (list.size()==1) {
			XWPFDocument word = WordUtils.getDocx(list.get(0));
			WordUtils.output(word, destPath);
			return;
		}
		
		XWPFDocument word1 = WordUtils.getDocx(list.get(0));
		
		for (int i=1; i<list.size(); i++) {
			XWPFDocument word2 = WordUtils.getDocx(list.get(i));
			word1 = mergeDocx(word1, word2, isPage);
		}
		
		WordUtils.output(word1, destPath);
	}
	
	/**
	 * 合并word(.docx后缀)
	 * @param word1
	 * @param word2
	 * @return
	 * @throws Exception 
	 */
	public XWPFDocument mergeDocx(XWPFDocument word1, XWPFDocument word2) throws Exception {
		return this.mergeDocx(word1, word2, true);
	}

	/**
	 * 合并word(.docx后缀)
	 * @param word1
	 * @param word2
	 * @param isPage       是否分页
	 * @return
	 * @throws Exception 
	 */
	public XWPFDocument mergeDocx(XWPFDocument word1, XWPFDocument word2, boolean isPage) throws Exception {
		if (isPage) {
			word1.createParagraph().createRun().addBreak(BreakType.PAGE);
		}
		
		CTBody body1 = word1.getDocument().getBody();
		CTBody body2 = word2.getDocument().getBody();
		
		List<XWPFPictureData> allPictures = word2.getAllPictures();
		// 记录图片合并前及合并后的ID
		Map<String, String> map = new HashMap<String, String>();
		for (XWPFPictureData picture : allPictures) {
			String before = word2.getRelationId(picture);
			// 将原文档中的图片加入到目标文档中
			String after = word1.addPictureData(picture.getData(), Document.PICTURE_TYPE_PNG);
			map.put(before, after);
		}
		
		// 合并内容
		this.appendBody(body1, body2, map);
		
		return word1;
	}

	/**
	 * 内容追加
	 * @param body1
	 * @param body2
	 * @param map
	 * @throws XmlException 
	 */
	private void appendBody(CTBody body1, CTBody body2, Map<String, String> map) throws XmlException {
		XmlOptions optionsOuter = new XmlOptions();
		optionsOuter.setSaveOuter();
		String appendString = body2.xmlText(optionsOuter);
		
		String srcString = body1.xmlText();
		String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
		String mainPart = srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
		String sufix = srcString.substring(srcString.lastIndexOf("<"));
		String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
		
		// 对xml字符串中图片ID进行替换
		if (map!=null && map.isEmpty()==false) {
			for (Map.Entry<String, String> set : map.entrySet()) {
				addPart = addPart.replace(set.getKey(), set.getValue());
			}
		}
		// 将两个文档的xml内容进行拼接
		CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
		body1.set(makeBody);
	}

}
