package cn.javaex.officejj.word.help;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import cn.javaex.officejj.common.util.PropertyHandler;

/**
 * Word模板增强处理。
 * 支持嵌套属性、表格行循环和简单条件段落，适合合同、报告、审批单等模板填充。
 *
 * @author 陈霓清
 */
public class WordTemplateHelper {
	private static final Pattern TABLE_LIST_PATTERN = Pattern.compile("\\$\\{([A-Za-z0-9_]+)\\.([^}]+)\\}");
	private static final Pattern IF_PATTERN = Pattern.compile("\\$\\{if\\s+([^}]+)\\}");
	private static final Pattern END_IF_PATTERN = Pattern.compile("\\$\\{/if\\}");

	/**
	 * 渲染增强模板。
	 * @param word Word文档
	 * @param param 参数
	 * @return
	 * @throws Exception
	 */
	public XWPFDocument render(XWPFDocument word, Map<String, Object> param) throws Exception {
		if (word==null) {
			throw new IllegalArgumentException("Word文档不能为空");
		}
		if (param==null || param.isEmpty()) {
			return word;
		}

		this.renderTableRows(word, param);
		this.renderConditionParagraphs(word, param);
		new ParagraphHelper().replaceParagraph(word, param);
		new TableHelper().replaceTable(word, param);
		return word;
	}

	/**
	 * 渲染表格行循环。
	 * 模板行中出现 ${list.name} 这类占位符时，如果参数 list 是 List，会按列表长度复制该模板行。
	 * @param word Word文档
	 * @param param 参数
	 * @throws Exception
	 */
	private void renderTableRows(XWPFDocument word, Map<String, Object> param) throws Exception {
		for (XWPFTable table : word.getTables()) {
			int rowIndex = 0;
			while (rowIndex<table.getRows().size()) {
				XWPFTableRow row = table.getRow(rowIndex);
				String rowText = row==null ? "" : row.getCtRow().xmlText();
				Matcher matcher = TABLE_LIST_PATTERN.matcher(rowText);
				if (!matcher.find()) {
					rowIndex++;
					continue;
				}

				String listKey = matcher.group(1);
				Object listObj = PropertyHandler.getValue(param, listKey);
				if (!(listObj instanceof List)) {
					rowIndex++;
					continue;
				}
				List<?> dataList = (List<?>) listObj;
				if (dataList.isEmpty()) {
					table.removeRow(rowIndex);
					continue;
				}

				for (int i=1; i<dataList.size(); i++) {
					XWPFTableRow newRow = table.insertNewTableRow(rowIndex + i);
					newRow.getCtRow().set(row.getCtRow().copy());
				}
				for (int i=0; i<dataList.size(); i++) {
					Map<String, Object> rowParam = new HashMap<String, Object>(param);
					rowParam.put(listKey, dataList.get(i));
					this.replaceRow(table.getRow(rowIndex + i), rowParam);
				}
				rowIndex += dataList.size();
			}
		}
	}

	/**
	 * 替换表格行中的占位符。
	 * @param row 表格行
	 * @param param 参数
	 * @throws Exception
	 */
	private void replaceRow(XWPFTableRow row, Map<String, Object> param) throws Exception {
		ParagraphHelper paragraphHelper = new ParagraphHelper();
		for (XWPFTableCell cell : row.getTableCells()) {
			for (XWPFParagraph paragraph : cell.getParagraphs()) {
				paragraphHelper.replaceParagraph(paragraph, param);
			}
		}
	}

	/**
	 * 渲染简单条件段落。
	 * 使用 ${if key} 和 ${/if} 包裹的段落块，当key为false、0、空字符串或null时整块移除。
	 * @param word Word文档
	 * @param param 参数
	 */
	private void renderConditionParagraphs(XWPFDocument word, Map<String, Object> param) {
		List<XWPFParagraph> paragraphs = word.getParagraphs();
		int index = 0;
		while (index<paragraphs.size()) {
			XWPFParagraph paragraph = paragraphs.get(index);
			Matcher startMatcher = IF_PATTERN.matcher(paragraph.getText());
			if (!startMatcher.find()) {
				index++;
				continue;
			}

			String key = startMatcher.group(1).trim();
			int endIndex = this.findEndIf(paragraphs, index + 1);
			if (endIndex<0) {
				index++;
				continue;
			}

			boolean show = PropertyHandler.isTrue(PropertyHandler.getValue(param, key));
			if (!show) {
				for (int i=endIndex; i>=index; i--) {
					word.removeBodyElement(word.getPosOfParagraph(paragraphs.get(i)));
				}
			} else {
				this.clearParagraph(paragraphs.get(endIndex));
				this.clearParagraph(paragraph);
				index = endIndex + 1;
			}
		}
	}

	/**
	 * 查找条件块结束段落。
	 * @param paragraphs 段落集合
	 * @param startIndex 起始索引
	 * @return
	 */
	private int findEndIf(List<XWPFParagraph> paragraphs, int startIndex) {
		for (int i=startIndex; i<paragraphs.size(); i++) {
			if (END_IF_PATTERN.matcher(paragraphs.get(i).getText()).find()) {
				return i;
			}
		}
		return -1;
	}

	/**
	 * 清空段落文本。
	 * @param paragraph 段落
	 */
	private void clearParagraph(XWPFParagraph paragraph) {
		for (XWPFRun run : paragraph.getRuns()) {
			run.setText("", 0);
		}
	}

	/**
	 * 设置书签文本。
	 * @param word Word文档
	 * @param bookmarkName 书签名称
	 * @param text 文本
	 */
	public void setBookmarkText(XWPFDocument word, String bookmarkName, String text) {
		for (XWPFParagraph paragraph : word.getParagraphs()) {
			for (CTBookmark bookmark : paragraph.getCTP().getBookmarkStartList()) {
				if (bookmarkName.equals(bookmark.getName())) {
					XWPFRun run = paragraph.createRun();
					run.setText(text==null ? "" : text);
					return;
				}
			}
		}
	}

	/**
	 * 给段落追加超链接。
	 * @param paragraph 段落
	 * @param text 显示文本
	 * @param url 链接地址
	 */
	public void addHyperlink(XWPFParagraph paragraph, String text, String url) throws Exception {
		String id = paragraph.getDocument().getPackagePart().addExternalRelationship(url,
				org.apache.poi.xwpf.usermodel.XWPFRelation.HYPERLINK.getRelation()).getId();
		CTP ctp = paragraph.getCTP();
		org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink hyperlink = ctp.addNewHyperlink();
		hyperlink.setId(id);
		CTR ctr = hyperlink.addNewR();
		XWPFHyperlinkRun run = new XWPFHyperlinkRun(hyperlink, ctr, paragraph);
		run.setText(text==null ? url : text);
		run.setColor("0563C1");
		run.setUnderline(org.apache.poi.xwpf.usermodel.UnderlinePatterns.SINGLE);
	}

	/**
	 * 插入简单目录占位字段。
	 * Word打开文档后可更新域得到目录。
	 * @param paragraph 段落
	 */
	public void addTocField(XWPFParagraph paragraph) {
		XWPFRun run = paragraph.createRun();
		run.getCTR().addNewFldChar().setFldCharType(org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType.BEGIN);
		run = paragraph.createRun();
		run.getCTR().addNewInstrText().setStringValue("TOC \\o \"1-3\" \\h \\z \\u");
		run = paragraph.createRun();
		run.getCTR().addNewFldChar().setFldCharType(org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType.SEPARATE);
		run = paragraph.createRun();
		run.setText("目录");
		run = paragraph.createRun();
		run.getCTR().addNewFldChar().setFldCharType(org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType.END);
	}
}
