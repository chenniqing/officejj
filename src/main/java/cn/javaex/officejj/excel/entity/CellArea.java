package cn.javaex.officejj.excel.entity;

/**
 * 单元格区域
 * 
 * @author 陈霓清
 */
public class CellArea {
	public Integer firstRow;     // 起始行（从0开始计算）
	public Integer lastRow;      // 终止行（从0开始计算）
	public Integer firstCol;     // 起始列（从0开始计算）
	public Integer lastCol;      // 终止列（从0开始计算）
	
	public CellArea() {
		super();
	}
	
	public CellArea(Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol) {
		super();
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
	}
	
	public Integer getFirstRow() {
		return firstRow;
	}
	public void setFirstRow(Integer firstRow) {
		this.firstRow = firstRow;
	}
	public Integer getLastRow() {
		return lastRow;
	}
	public void setLastRow(Integer lastRow) {
		this.lastRow = lastRow;
	}
	public Integer getFirstCol() {
		return firstCol;
	}
	public void setFirstCol(Integer firstCol) {
		this.firstCol = firstCol;
	}
	public Integer getLastCol() {
		return lastCol;
	}
	public void setLastCol(Integer lastCol) {
		this.lastCol = lastCol;
	}
	
}
