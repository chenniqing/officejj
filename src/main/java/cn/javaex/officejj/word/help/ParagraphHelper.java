package cn.javaex.officejj.word.help;

import java.io.InputStream;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import cn.javaex.officejj.common.entity.Font;
import cn.javaex.officejj.common.entity.Picture;
import cn.javaex.officejj.common.entity.Table;
import cn.javaex.officejj.common.util.ImageHandler;

/**
 * 段落
 * 
 * @author 陈霓清
 */
public class ParagraphHelper extends Helper {

	/**
	 * 替换段落变量
	 * @param word
	 * @param param
	 * @throws Exception
	 */
	public void replaceParagraph(XWPFDocument word, Map<String, Object> param) throws Exception {
		List<XWPFParagraph> paragraphList = word.getParagraphs();
		
		if (paragraphList!=null && paragraphList.isEmpty()==false) {
			for (XWPFParagraph paragraph : paragraphList) {
				this.replaceParagraph(paragraph, param);
			}
		}
	}
	
	/**
	 * 替换段落变量
	 * @param paragraph
	 * @param param
	 * @throws Exception
	 */
	public void replaceParagraph(XWPFParagraph paragraph, Map<String, Object> param) throws Exception {
		String tempString = "";
		Set<XWPFRun> runSet = new HashSet<XWPFRun>();
		char lastChar = ' ';
		List<XWPFRun> runList = paragraph.getRuns();
		for (XWPFRun run : runList) {
			String text = run.getText(0);
			if (text==null) {
				continue;
			}
			
			run.setText(text, 0);
			for (int i=0; i<text.length(); i++) {
				char ch = text.charAt(i);
				if (ch=='$') {
					runSet = new HashSet<XWPFRun>();
					runSet.add(run);
					tempString = text;
				}
				else if (ch=='{') {
					if (lastChar=='$') {
						if (runSet.contains(run)) {
							
						} else {
							runSet.add(run);
							tempString = tempString + text;
						}
					} else {
						runSet = new HashSet<XWPFRun>();
						tempString = "";
					}
				}
				else if (ch=='}') {
					if (tempString!=null && tempString.contains("${")) {
						if (runSet.contains(run)) {
							
						} else {
							runSet.add(run);
							tempString = tempString + text;
						}
					} else {
						runSet = new HashSet<XWPFRun>();
						tempString = ""; 
					}
					if (runSet.size()>0) {
						String replaceContent = this.replaceContent(tempString, param, run);
						if (!replaceContent.equals(tempString)) {
							int index = 0;
							XWPFRun xwpfRun = null;
							for (XWPFRun tempRun : runSet) {
								tempRun.setText("", 0);
								if (index==0) {
									xwpfRun = tempRun;
								}
								index++;
							}
							xwpfRun.setText(replaceContent, 0);
						}
						runSet = new HashSet<XWPFRun>();
						tempString = ""; 
					}
				}
				else {
					if (runSet.size()<=0) {
						continue;
					}
					if (runSet.contains(run)) {
						continue;
					}
					runSet.add(run);
					tempString = tempString + text;
				}
				
				lastChar = ch;
			}
		}
	}
	
	/**
	 * 替换内容
	 * @param text
	 * @param param
	 * @param run
	 * @return
	 * @throws Exception 
	 */
	private String replaceContent(String text, Map<String, Object> param, XWPFRun run) throws Exception {
		ImageHelper imageHelper = new ImageHelper();
		RunHelper runHelper = new RunHelper();
		
		if (text==null || text.length()==0) {
			return text;
		}
		
		// 去除前后的 ${ 和 }
		String key = super.getPlaceholder(text);
		if (key.length()==0) {
			return text;
		}
		String replaceKey = "${" + key + "}";
		
		Object obj = param.get(key);
		if (obj==null) {
			return text;
		}
		
		// 文本替换
		if (obj instanceof String) {
			String str = this.replaceBr(obj.toString());
			if (str.contains("<br/>")) {
				text = text.replace(replaceKey, "");
				runHelper.setWrapText(run, str);
			} else {
				text = text.replace(replaceKey, str);
			}
		}
		// 文本替换（带字体样式）
		else if (obj instanceof Font) {
			Font font = (Font) obj;
			
			// 设置字体样式
			runHelper.setFontStyle(run, font);
			// 设置文本
			String str = super.replaceBr(font.getText());
			if (str.contains("<br/>")) {
				text = text.replace(replaceKey, "");
				runHelper.setWrapText(run, str);
			} else {
				text = text.replace(replaceKey, str);
			}
		}
		// 图片替换
		else if (obj instanceof Picture) {
			InputStream in = null;
			try {
				Picture picture = (Picture) obj;
				
				double width = picture.getWidth();
				double height = picture.getHeight();
				String imgUrl = picture.getUrl();
				String imgType = imgUrl.substring(imgUrl.lastIndexOf(".") + 1);
				int imageType = imageHelper.getImageType(imgType);
				
				// 获得图片流
				in = ImageHandler.getImageStream(imgUrl);
				if (in!=null) {
					run.addPicture(in, imageType, null, Units.toEMU(width), Units.toEMU(height));
				}
				
				// 图片描述
				if (picture.getDescription()!=null) {
					run.addBreak();
					run.setText(picture.getDescription());
				}
				
				text = text.replace(replaceKey, "");
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				IOUtils.closeQuietly(in);
			}
		}
		// 插入表格数据
		else if (obj instanceof Table) {
			throw new Exception(key);
		}
		// 数字之类的，直接强转为字符串
		else {
			String str = this.replaceBr(obj.toString());
			if (str.contains("<br/>")) {
				text = text.replace(replaceKey, "");
				runHelper.setWrapText(run, str);
			} else {
				text = text.replace(replaceKey, str);
			}
		}
		
		return text;
	}
	
	/**
	 * 设置指定变量为指定值
	 * @param paragraph
	 * @param value
	 */
	public void replaceParagraph(XWPFParagraph paragraph, String value) {
		List<XWPFRun> runList = paragraph.getRuns();
		for (XWPFRun run : runList) {
			run.setText(value, 0);
		}
	}
	
}
