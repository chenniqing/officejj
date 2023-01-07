package cn.javaex.officejj.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 自定义样式
 * 
 * @author 陈霓清
 */
public class DefaultCellStyle implements ICellStyle {

	/**
	 * 创建标题样式
	 */
	@Override
	public CellStyle createTitleStyle(Workbook wb) {
		// 设置字体样式
		CellStyle cellStyle = wb.createCellStyle();
		// 水平对齐方式（居中）
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		// 垂直对齐方式（居中）
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		// 字体
		Font font = wb.createFont();
		font.setFontName("等线");
		font.setBold(true);
		font.setFontHeightInPoints((short) 16);
		cellStyle.setFont(font);
		
		return cellStyle;
	}
	
	/**
	 * 创建头部样式
	 */
	@Override
	public CellStyle createHeaderStyle(Workbook wb) {
		// 设置字体样式
		CellStyle cellStyle = wb.createCellStyle();
		// 水平对齐方式（居中）
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		// 垂直对齐方式（居中）
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		// 字体
		Font font = wb.createFont();
		font.setFontName("等线");
		cellStyle.setFont(font);
		
		return cellStyle;
	}

	/**
	 * 创建数据样式
	 */
	@Override
	public CellStyle createDataStyle(Workbook wb) {
		// 设置字体样式
		CellStyle cellStyle = wb.createCellStyle();
		// 水平对齐方式（居中）
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		// 垂直对齐方式（居中）
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		// 自动换行
		cellStyle.setWrapText(true);
		// 字体
		Font font = wb.createFont();
		font.setFontName("等线");
		cellStyle.setFont(font);
		
		return cellStyle;
	}
	
}
