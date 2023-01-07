package cn.javaex.officejj.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 单元格样式接口
 * 
 * @author 陈霓清
 */
public interface ICellStyle {
	
	/**
	 * 创建头部样式
	 * @param wk
	 * @return 
	 */
	CellStyle createTitleStyle(Workbook wk);
	
	/**
	 * 创建头部样式
	 * @param wk
	 * @return 
	 */
	CellStyle createHeaderStyle(Workbook wk);
	
	/**
	 * 创建数据样式
	 * @param wk
	 * @return 
	 */
	CellStyle createDataStyle(Workbook wk);
	
}
