package cn.javaex.officejj.excel.entity;

/**
 * 纵向合并
 * 
 * @author 陈霓清
 */
public class VerticalMerge extends Merge {

	public VerticalMerge() {
		
	}

	/**
	 * @param firstRow
	 * @param lastRow
	 * @param firstCol
	 * @param lastCol
	 */
	public VerticalMerge(Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol) {
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
	}
	
}
