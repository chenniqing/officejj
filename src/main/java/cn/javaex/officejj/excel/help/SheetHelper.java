package cn.javaex.officejj.excel.help;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

import cn.javaex.officejj.excel.entity.ExcelSetting;

/**
 * Sheet操作
 * 
 * @author 陈霓清
 */
public class SheetHelper {
	
	/** 默认sheet页名称 */
	public static final String SHEET_NAME = "Sheet1";
	/** 行高基数 */
	public static final int BASE_ROW_HEIGHT = 20;
	/** 列宽基数 */
	public static final int BASE_COLUMN_WIDTH = 256;
	
	// 存储值替换
	public Map<String, Object> replaceMap = new HashMap<String, Object>();
	// 存储格式化
	public Map<String, Object> formatMap = new HashMap<String, Object>();
	// 存储合并多个单元格数据的成员变量
	public Map<String, String> skipMap = new HashMap<String, String>();
	
	/**
	 * 创建Header
	 * @param sheet
	 * @param clazz
	 * @param title
	 * @throws Exception
	 */
	public void setHeader(Sheet sheet, Class<?> clazz, String title) throws Exception {
		
	}
	
	/**
	 * 设置基本属性
	 * @param sheet
	 * @param clazz
	 */
	public void setBasicData(Sheet sheet, Class<?> clazz) {
		
	}
	
	/**
	 * 根据注解创建内容
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param title
	 * @throws Exception 
	 */
	public void write(Sheet sheet, Class<?> clazz, List<?> list, String title) throws Exception {
		
	}

	/**
	 * 根据注解创建内容
	 * 多线程
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @throws Exception
	 */
	public void writeByThreads(Sheet sheet, Class<?> clazz, List<?> list) throws Exception {
		
	}
	
	/**
	 * 根据设置类创建内容
	 * @param sheet
	 * @param excelSetting
	 */
	public void write(Sheet sheet, ExcelSetting excelSetting) {
		
	}
	
	/**
	 * 根据模板占位符替换内容
	 * @param sheet
	 * @param param
	 */
	public void write(Sheet sheet, Map<String, Object> param) {
		
	}

	/**
	 * 读取sheet
	 * @param <T>
	 * @param sheet
	 * @param clazz     自定义实体类
	 * @param rowNum    从第几行开始读取（从0开始计算）
	 * @return
	 * @throws Exception 
	 */
	public <T> List<T> read(Sheet sheet, Class<T> clazz, int rowNum) throws Exception {
		return null;
	}
	
	/**
	 * 设置下拉选项
	 * @param sheet
	 * @param colNum          第几个列（从0开始计算）
	 * @param startRow        第几个行设置开始（从0开始计算）
	 * @param endRow          第几个行设置结束（从0开始计算）
	 * @param selectDataArr   下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public void setSelect(Sheet sheet, int colIndex, int startRowIndex, int endRowIndex, String[] selectDataArr) {
		if (selectDataArr==null || selectDataArr.length==0) {
			return;
		}
		
		// 获取单元格样式
		CellStyle cellStyle = null;
		try {
			// 获取第一个单元格的样式，用于继承
			cellStyle = sheet.getRow(startRowIndex).getCell(colIndex).getCellStyle();
		} catch (Exception e) {
			// 如果没有该单元格存在，则使用默认的样式
			cellStyle = sheet.getWorkbook().createCellStyle();
		}
		// 自动换行
		cellStyle.setWrapText(true);
		
		Row row = null;
		Cell cell = null;
		
		for (int i=startRowIndex; i<=endRowIndex; i++) {
			row = sheet.getRow(i);
			if (row==null) {
				cell = sheet.createRow(i).createCell(colIndex);
			} else {
				cell = row.getCell(colIndex);
				if (cell==null) {
					cell = row.createCell(colIndex);
				}
			}
			
			cell.setCellStyle(cellStyle);
		}
		
		// 下拉的数据、起始行、终止行、起始列、终止列
		CellRangeAddressList addressList = new CellRangeAddressList(startRowIndex, endRowIndex, colIndex, colIndex);
		
		// 生成下拉框内容
		DataValidationHelper helper = sheet.getDataValidationHelper();
		DataValidationConstraint constraint = helper.createExplicitListConstraint(selectDataArr); 
		DataValidation dataValidation = helper.createValidation(constraint, addressList);
		
		// 设置数据有效性
		sheet.addValidationData(dataValidation);
	}

	/**
	 * 设置合并
	 * @param wb
	 * @param firstRow    起始行（从0开始计算）
	 * @param lastRow     终止行（从0开始计算）
	 * @param firstCol    起始列（从0开始计算）
	 * @param lastCol     终止列（从0开始计算）
	 */
	public void setMerge(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		sheet.addMergedRegion(cellRangeAddress);
	}
	
	/**
	 * 插入行
	 * @param sheet
	 * @param startRow
	 * @param rows
	 */
	public static void insertRow(Sheet sheet, int startRow, int rows) {
		if (rows==0) {
			return;
		}
		
		// 解决list占位符在最后一行时报错的BUG
		if ((startRow + 1) > sheet.getLastRowNum()) {
			sheet.createRow(startRow + 2);
		}
		
		/**
		 * startRow                  从下标为startRow的行开始移动
		 * endRow                    到下标为endRow的行结束移动
		 * n                         有多少行需要移动
		 * copyRowHeight             是否复制行高
		 * resetOriginalRowHeight    是否将原始行的高度设置为默认
		 */
		sheet.shiftRows(startRow + 1, sheet.getLastRowNum(), rows, true, false);
		
		RowHelper rowHelper = new RowHelper();
		
		for (int i=0; i<rows; i++) {
			Row sourceRow = null;
			Row targetRow = null;
			
			sourceRow = sheet.getRow(startRow);
			targetRow = sheet.createRow(++startRow);
			
			rowHelper.copyRow(sheet, sourceRow, targetRow);
		}
	}

}
