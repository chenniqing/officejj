package cn.javaex.officejj.word.help;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.namespace.QName;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import cn.javaex.officejj.word.WordUtils;

/**
 * 合并Word
 *
 * @author 陈霓清
 */
public class MergeHelper {

	private static final QName BODY_SECT_PR = new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "sectPr");

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
			try {
				word1 = mergeDocx(word1, word2, isPage);
			} finally {
				if (word2!=null) {
					word2.close();
				}
			}
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
		try (XmlCursor targetCursor = body1.newCursor(); XmlCursor sourceCursor = body2.newCursor()) {
			this.moveToBodyInsertPosition(targetCursor);
			if (sourceCursor.toFirstChild()) {
				do {
					if (BODY_SECT_PR.equals(sourceCursor.getName())) {
						continue;
					}
					this.copyBodyChild(sourceCursor, targetCursor, map);
				} while (sourceCursor.toNextSibling());
			}
		}
	}

	/**
	 * 定位到正文追加位置。
	 * <p>Word 的 sectPr 必须保留在 body 末尾，追加内容时需要插入到 sectPr 之前。</p>
	 * @param targetCursor 目标正文游标
	 */
	private void moveToBodyInsertPosition(XmlCursor targetCursor) {
		if (targetCursor.toChild(BODY_SECT_PR)) {
			return;
		}
		targetCursor.toEndToken();
	}

	/**
	 * 复制正文子节点。
	 * <p>逐个复制完整节点可以保留命名空间声明，避免字符串裁剪后出现 main 前缀未绑定的问题。</p>
	 * @param sourceCursor 源正文子节点游标
	 * @param targetCursor 目标正文插入游标
	 * @param map 图片关系 ID 替换表
	 * @throws XmlException XML 解析失败时抛出
	 */
	private void copyBodyChild(XmlCursor sourceCursor, XmlCursor targetCursor, Map<String, String> map) throws XmlException {
		String childXml = this.replacePictureRelationIds(sourceCursor.xmlText(), map);
		XmlObject child = XmlObject.Factory.parse(childXml);
		try (XmlCursor childCursor = child.newCursor()) {
			if (childCursor.toFirstChild()) {
				childCursor.copyXml(targetCursor);
			}
		}
	}

	/**
	 * 替换复制内容中的图片关系 ID。
	 * @param xml 复制节点 XML
	 * @param map 图片关系 ID 替换表
	 * @return 替换后的 XML
	 */
	private String replacePictureRelationIds(String xml, Map<String, String> map) {
		if (map==null || map.isEmpty()) {
			return xml;
		}
		String replacedXml = xml;
		for (Map.Entry<String, String> set : map.entrySet()) {
			replacedXml = replacedXml.replace(set.getKey(), set.getValue());
		}
		return replacedXml;
	}

}
