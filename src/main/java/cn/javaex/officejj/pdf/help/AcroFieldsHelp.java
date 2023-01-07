package cn.javaex.officejj.pdf.help;

import java.io.IOException;
import java.util.Map;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfStamper;

import cn.javaex.officejj.common.entity.Font;
import cn.javaex.officejj.common.entity.Picture;
import cn.javaex.officejj.common.entity.RGB;

/**
 * 表单填充
 * 
 * @author 陈霓清
 */
public class AcroFieldsHelp extends Helper {
	/**
	 * 替换占位符内容
	 * @param form
	 * @param stamper
	 * @param param
	 * @return
	 * @throws IOException
	 * @throws DocumentException
	 */
	public AcroFields replaceContent(AcroFields form, PdfStamper stamper, Map<String, Object> param) throws IOException, DocumentException {
		if (param==null || param.size()==0) {
			return form;
		}
		
		for (Map.Entry<String, Object> entry : param.entrySet()) {
			String key = entry.getKey();
			Object value = entry.getValue();
			
			if (value==null) {
				value = "";
			}
			
			// 文本替换
			if (value instanceof String) {
				form.setField(key, (String) value);
			}
			// 自定义字体样式
			else if (value instanceof Font) {
				Font font = (Font) value;
				
				if (font.getFontFamily()!=null) {
					String path = super.getRealPath(font.getFontFamily());
					BaseFont baseFont = BaseFont.createFont(path, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
					form.setFieldProperty(key, "textfont", baseFont, null);
				}
				if (font.getColor()!=null) {
					RGB rgb = new RGB(font.getColor());
					BaseColor baseColor = new BaseColor(rgb.getRed(), rgb.getGreen(), rgb.getBlue());
					form.setFieldProperty(key, "textcolor", baseColor, null);
				}
				if (font.getFontSize()!=null) {
					form.setFieldProperty(key, "textsize", font.getFontSize().floatValue(), null);
				}
				
				form.setField(key, font.getText());
			}
			// 图片替换
			else if (value instanceof Picture) {
				Picture picture = (Picture) value;
				
				// 获取所在页和坐标，左下角为起点
				int pageNo = form.getFieldPositions(key).get(0).page;
				Rectangle signRect = form.getFieldPositions(key).get(0).position;
				float x = signRect.getLeft();
				float y = signRect.getBottom();
				
				// 读取图片
				Image image = Image.getInstance(picture.getUrl());
				// 获取操作的页面
				PdfContentByte under = stamper.getOverContent(pageNo);
				// 设置图片大小
				if (picture.getWidth()==null || picture.getHeight()==null) {
					image.scaleToFit(signRect.getWidth(), signRect.getHeight());    // 根据域的大小缩放图片（图片大小自适应）
				} else {
					double width = picture.getWidth();
					double height = picture.getHeight();
					image.scaleAbsolute((float) width, (float) height);    // 指定图片大小
				}
				// 添加图片
				image.setAbsolutePosition(x, y);
				under.addImage(image);
			}
			// 数字之类的直接转字符串
			else {
				form.setField(key, value.toString());
			}
		}
		
		return form;
	}
	
}
