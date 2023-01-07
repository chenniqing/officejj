package cn.javaex.officejj.common.entity;

import java.io.ByteArrayOutputStream;
import java.io.ObjectOutputStream;
import java.io.Serializable;

/**
 * 字体样式
 * 
 * @author 陈霓清
 */
public class Font implements Serializable {
	
	private static final long serialVersionUID = 1L;
	
	private String text;             // 文本内容
	private String color;            // 颜色：RGB，例如：FF0000
	private String fontFamily;       // 字体
	private Integer fontSize;        // 字体大小
	private boolean bold;            // 粗体
	private boolean italic;          // 斜体
	private boolean strike;          // 删除线

	/**
	 * 得到文本内容
	 * @return
	 */
	public String getText() {
		return text;
	}
	/**
	 * 设置文本内容
	 * @param text
	 */
	public void setText(String text) {
		this.text = text;
	}

	/**
	 * 得到颜色
	 * @return
	 */
	public String getColor() {
		return color;
	}
	/**
	 * 设置颜色
	 * @param color
	 */
	public void setColor(String color) {
		this.color = color;
	}

	/**
	 * 得到字体
	 * @return
	 */
	public String getFontFamily() {
		return fontFamily;
	}
	/**
	 * 设置字体
	 * @param fontFamily
	 */
	public void setFontFamily(String fontFamily) {
		this.fontFamily = fontFamily;
	}

	/**
	 * 得到字体大小
	 * @return
	 */
	public Integer getFontSize() {
		return fontSize;
	}
	/**
	 * 设置字体大小
	 * @param fontSize
	 */
	public void setFontSize(Integer fontSize) {
		this.fontSize = fontSize;
	}

	/**
	 * 得到粗体
	 * @return
	 */
	public boolean getBold() {
		return bold;
	}
	/**
	 * 设置粗体
	 * @param bold
	 */
	public void setBold(boolean bold) {
		this.bold = bold;
	}

	/**
	 * 得到斜体
	 * @return
	 */
	public boolean getItalic() {
		return italic;
	}
	/**
	 * 设置斜体
	 * @param italic
	 */
	public void setItalic(boolean italic) {
		this.italic = italic;
	}

	/**
	 * 得到删除线
	 * @return
	 */
	public boolean getStrike() {
		return strike;
	}
	/**
	 * 设置删除线
	 * @param strike
	 */
	public void setStrike(boolean strike) {
		this.strike = strike;
	}
	
	/**
	 * 对象序列化
	 */
	@Override
	public String toString() {
		String objStr = "";
		
		try {
			ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
			ObjectOutputStream out = new ObjectOutputStream(byteArrayOutputStream);
			out.writeObject(this);
			objStr = byteArrayOutputStream.toString("ISO-8859-1");
			out.close();
			byteArrayOutputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return objStr;
	}

}
