package cn.javaex.officejj.excel.entity;

/**
 * 横向合并
 * 
 * @author 陈霓清
 */
public class TransversalMerge extends Merge {

	public TransversalMerge() {
		
	}

	/**
	 * @param firstRow
	 * @param lastRow
	 * @param firstCol
	 * @param lastCol
	 */
	public TransversalMerge(Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol) {
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
	}
	
}
