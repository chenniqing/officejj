package cn.javaex.officejj.pdf.help;

import java.io.FileOutputStream;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGState;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

import cn.javaex.officejj.common.entity.Font;

/**
 * Pdf
 * 
 * @author 陈霓清
 */
public class PdfHelper extends Helper {

	/**
	 * 设置水印
	 * @param reader
	 * @param obj       水印内容，可以是纯英文文字，如果是中文的话，必须使用cn.javaex.officejj.common.entity.Font
	 * @param destPath
	 */
	public void setWatermark(PdfReader reader, Object obj, String destPath) throws Exception {
		PdfStamper stamper = new PdfStamper(reader, new FileOutputStream(destPath));
		int total = reader.getNumberOfPages() + 1;
		
		// 获取页面宽高
		Document document = new Document(reader.getPageSize(1)); 
		float width = document.getPageSize().getWidth();
		float height = document.getPageSize().getHeight();
		
		PdfContentByte content;
		
		// 设置字体
		BaseFont baseFont = null;
		if (obj instanceof Font) {
			Font font = (Font) obj;
			String path = super.getRealPath(font.getFontFamily());
			baseFont = BaseFont.createFont(path, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
			
			// 循环对每页插入水印
			for (int i = 1; i < total; i++) {
				// 水印的起始
				content = stamper.getUnderContent(i);
				// 开始
				content.beginText();
				content.setColorFill(BaseColor.GRAY);
				// 设置字体及字号
				content.setFontAndSize(baseFont, 38);
				if (font.getFontSize()!=null) {
					content.setFontAndSize(baseFont, font.getFontSize());
				}
				// 设置透明度
				PdfGState pdfGState = new PdfGState();
				pdfGState.setFillOpacity(0.3f);
				content.setGState(pdfGState);
				// 开始写入水印
				content.showTextAligned(Element.ALIGN_LEFT, font.getText(), width/3, height/3, 40);
				content.endText();
			}
		} else {
			baseFont = BaseFont.createFont();
			
			// 循环对每页插入水印
			for (int i = 1; i < total; i++) {
				// 水印的起始
				content = stamper.getUnderContent(i);
				// 开始
				content.beginText();
				content.setColorFill(BaseColor.GRAY);
				// 设置字体及字号
				content.setFontAndSize(baseFont, 38);
				// 设置透明度
				PdfGState pdfGState = new PdfGState();
				pdfGState.setFillOpacity(0.3f);
				content.setGState(pdfGState);
				// 开始写入水印
				content.showTextAligned(Element.ALIGN_LEFT, (String) obj, width/3, height/3, 40);
				content.endText();
			}
		}
		
		stamper.close();
	}

}
