package cn.javaex.officejj.excel.entity;

/**
 * 打印区域
 * 
 * @author 陈霓清
 * @Date 2026年2月14日
 */
public class SheetPrintArea {
	public int firstRow;          // 起始行，1-based，例如1表示Excel的第1行
	public int lastRow;           // 终止行，1-based，包含
	public String firstColumn;    // 起始列，如"A"
	public String lastColumn;     // 终止列，如"F"
	
	public SheetPrintArea() {
		super();
	}

	public SheetPrintArea(int firstRow, int lastRow, String firstColumn, String lastColumn) {
		super();
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstColumn = firstColumn;
		this.lastColumn = lastColumn;
	}
	
	public int getFirstRow() {
		return firstRow;
	}
	public void setFirstRow(int firstRow) {
		this.firstRow = firstRow;
	}
	public int getLastRow() {
		return lastRow;
	}
	public void setLastRow(int lastRow) {
		this.lastRow = lastRow;
	}
	public String getFirstColumn() {
		return firstColumn;
	}
	public void setFirstColumn(String firstColumn) {
		this.firstColumn = firstColumn;
	}
	public String getLastColumn() {
		return lastColumn;
	}
	public void setLastColumn(String lastColumn) {
		this.lastColumn = lastColumn;
	}
	
}
