package cn.javaex.officejj.word.help;

import java.io.InputStream;

import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import cn.javaex.officejj.common.entity.Font;
import cn.javaex.officejj.common.entity.Picture;
import cn.javaex.officejj.common.util.ImageHandler;

/**
 * 一段文本
 * 
 * @author 陈霓清
 */
public class RunHelper extends Helper {

	/**
	 * 设置字体样式
	 * @param run
	 * @param font
	 */
	public void setFontStyle(XWPFRun run, Font font) {
		if (font.getColor()!=null) {
			run.setColor(font.getColor());
		}
		if (font.getFontFamily()!=null) {
			run.setFontFamily(font.getFontFamily());
		}
		if (font.getFontSize()!=null && font.getFontSize()>0) {
			run.setFontSize(font.getFontSize());
		}
		if (font.getBold()) {
			run.setBold(true);
		}
		if (font.getItalic()) {
			run.setItalic(true);
		}
		if (font.getStrike()) {
			run.setStrikeThrough(true);
		}
	}
	
	/**
	 * 设置值
	 * @param run
	 * @param obj
	 */
	public void setValue(XWPFRun run, Object obj) {
		if (obj==null) {
			obj = "";
		}
		
		// 文本
		if (obj instanceof String) {
			run.setText((String) obj);
		}
		// 文本替换
		else if (obj instanceof Font) {
			Font font = (Font) obj;
			
			// 设置字体样式和文本
			this.setFontStyle(run, font);
			this.setText(run, font.getText());
		}
		// 图片
		else if (obj instanceof Picture) {
			ImageHelper imageHelper = new ImageHelper();
			
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
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				IOUtils.closeQuietly(in);
			}
		}
		// 数字之类的，直接强转为字符串
		else {
			run.setText(obj.toString());
		}
	}
	
	/**
	 * 设置文本
	 * @param run
	 * @param text
	 */
	public void setText(XWPFRun run, String text) {
		text = super.replaceBr(text);
		if (text.contains("<br/>")) {
			this.setWrapText(run, text);
		} else {
			run.setText(text);
		}
	}
	
	/**
	 * 设置换行文本
	 * @param run
	 * @param text
	 */
	public void setWrapText(XWPFRun run, String text) {
		String[] arr = text.split("<br/>");
		for (int n=0; n<arr.length; n++) {
			if (n==0) {
				run.setText(arr[n]);
			} else {
				run.addBreak();
				run.setText(arr[n]);
			}
		}
	}

}
