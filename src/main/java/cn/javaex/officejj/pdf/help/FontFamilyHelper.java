package cn.javaex.officejj.pdf.help;

import java.io.File;
import java.net.URL;

import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.tool.xml.XMLWorkerFontProvider;

import cn.javaex.officejj.common.util.PathHandler;

/**
 * 自定义字体
 * 
 * @author 陈霓清
 */
public class FontFamilyHelper extends XMLWorkerFontProvider {
	private String fontFamily;
	
	public FontFamilyHelper(String fontFamily) {
		this.fontFamily = fontFamily;
	}
	
	public String getFontFamily() {
		return fontFamily;
	}

	public void setFontFamily(String fontFamily) {
		this.fontFamily = fontFamily;
	}

	@Override
	public Font getFont(final String fontname, String encoding, float size, final int style) {
		try {
			String path = this.fontFamily;
			
			// resources文件夹下的字体
			if (path.startsWith("resources:")) {
				path = path.replace("resources:", "");
				if (path.startsWith("/")) {
					path = path.substring(1);
				}
				
				URL fontPath = Thread.currentThread().getContextClassLoader().getResource(path);
				path = fontPath + "";
				if (path.endsWith(".ttc")) {
					path = path + ",1";
				}
			} else {
				boolean absolutePath = PathHandler.isAbsolutePath(path);
				
				if (!absolutePath) {
					String projectPath = PathHandler.getProjectPath();
					path = projectPath + File.separator + path;
				}
			}
			
			if (path.endsWith(".ttc")) {
				path = path + ",1";
			}
			
			BaseFont bfChinese = BaseFont.createFont(path, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
			return new Font(bfChinese, size, style);
		} catch (Exception e) {
			
		}
		
		return super.getFont(fontname, encoding, size, style);
	}
}
