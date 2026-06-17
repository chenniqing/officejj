package cn.javaex.officejj.excel.entity;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel SAX读取到的一行数据。
 *
 * @author 陈霓清
 */
public class ExcelSaxRow {
	private int sheetNum;                    // 第几个Sheet，从1开始
	private String sheetName;                // Sheet名称
	private int rowNum;                      // Excel行号，从1开始
	private List<String> cellList = new ArrayList<String>(); // 单元格文本，下标从0开始

	public int getSheetNum() {
		return sheetNum;
	}

	public void setSheetNum(int sheetNum) {
		this.sheetNum = sheetNum;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public int getRowNum() {
		return rowNum;
	}

	public void setRowNum(int rowNum) {
		this.rowNum = rowNum;
	}

	public List<String> getCellList() {
		return cellList;
	}

	public void setCellList(List<String> cellList) {
		this.cellList = cellList;
	}

	/**
	 * 读取指定列文本。
	 * @param colNum 第几列，从1开始
	 * @return
	 */
	public String getCellValue(int colNum) {
		if (colNum<=0 || cellList==null || colNum>cellList.size()) {
			return "";
		}
		String value = cellList.get(colNum - 1);
		return value==null ? "" : value;
	}
}
