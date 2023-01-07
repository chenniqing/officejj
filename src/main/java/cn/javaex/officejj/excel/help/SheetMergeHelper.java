package cn.javaex.officejj.excel.help;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import cn.javaex.officejj.excel.entity.TransversalMerge;
import cn.javaex.officejj.excel.entity.VerticalMerge;

/**
 * 合并
 * 
 * @author 陈霓清
 */
public class SheetMergeHelper extends SheetHelper {
	
	private String transverseVlaue = null;                // 横向合并值
	private String verticalVlaue = null;                  // 纵向合并值
	private TransversalMerge transversalMerge = null;     // 横向合并类
	private VerticalMerge verticalMerge = null;           // 纵向合并类
	
	/**
	 * 设置表头合并
	 * @param sheet
	 * @param rowIndex
	 * @param headerRows
	 */
	public void setHeaderMerge(Sheet sheet, int rowIndex, int headerRows) {
		// 设置横向合并
		this.setTransverseMerge(sheet, rowIndex, headerRows);
		
		// 设置纵向合并
		this.setVerticalMerge(sheet, rowIndex, headerRows);
	}

	/**
	 * 设置横向合并
	 * @param sheet
	 * @param rowIndex
	 * @param headerRows
	 */
	private void setTransverseMerge(Sheet sheet, int rowIndex, int headerRows) {
		for (int i=0; i<headerRows; i++) {
			Row row = sheet.getRow(i);
			int lastCellNum = row.getLastCellNum();    // 一共多少列
			
			for (int j=0; j<lastCellNum; j++) {
				Cell cell = row.getCell(j);
				String cellValue = cell.getRichStringCellValue().getString();
				
				if (transverseVlaue==null) {
					transverseVlaue = cellValue;
				} else {
					if (transverseVlaue.equals(cellValue)) {
						if (transversalMerge==null) {
							transversalMerge = new TransversalMerge(i, i, j-1, j);
						} else {
							transversalMerge = new TransversalMerge(i, i, transversalMerge.getFirstCol(), j);
						}
					} else {
						if (transversalMerge!=null) {
							super.setMerge(sheet, transversalMerge.getFirstRow(), transversalMerge.getLastRow(), transversalMerge.getFirstCol(), transversalMerge.getLastCol());
							transverseVlaue = null;
							transversalMerge = null;
						}
						
						transverseVlaue = cellValue;
					}
				}
				
				// 每一行遍历完成后设置横向合并
				if (j == (lastCellNum-1)) {
					transverseVlaue = null;
					
					if (transversalMerge!=null) {
						super.setMerge(sheet, transversalMerge.getFirstRow(), transversalMerge.getLastRow(), transversalMerge.getFirstCol(), transversalMerge.getLastCol());
						transverseVlaue = null;
						transversalMerge = null;
					}
				}
			}
		}
		
		// 最后一行遍历完成后设置横向合并
		if (transversalMerge!=null) {
			super.setMerge(sheet, transversalMerge.getFirstRow(), transversalMerge.getLastRow(), transversalMerge.getFirstCol(), transversalMerge.getLastCol());
			transverseVlaue = null;
			transversalMerge = null;
		}
	}
	
	/**
	 * 设置纵向合并
	 * @param sheet
	 * @param rowIndex
	 * @param headerRows
	 */
	private void setVerticalMerge(Sheet sheet, int rowIndex, int headerRows) {
		Row row = sheet.getRow(rowIndex);
		int lastCellNum = row.getLastCellNum();    // 一共多少列
		
		for (int i=0; i<lastCellNum; i++) {
			for (int j=0; j<headerRows; j++) {
				int rowIndexNew = rowIndex + j;
				row = sheet.getRow(rowIndexNew);
				
				Cell cell = row.getCell(i);
				String cellValue = cell.getRichStringCellValue().getString();
				
				if (verticalVlaue==null) {
					verticalVlaue = cellValue;
				} else {
					if (verticalVlaue.equals(cellValue)) {
						if (verticalMerge==null) {
							verticalMerge = new VerticalMerge(rowIndexNew-1, rowIndexNew, i, i);
						} else {
							verticalMerge = new VerticalMerge(verticalMerge.getFirstRow(), rowIndexNew, i, i);
						}
					} else {
						if (verticalMerge!=null) {
							super.setMerge(sheet, verticalMerge.getFirstRow(), verticalMerge.getLastRow(), verticalMerge.getFirstCol(), verticalMerge.getLastCol());
							verticalVlaue = null;
							verticalMerge = null;
						}
						
						verticalVlaue = cellValue;
					}
				}
				
				// 每一列遍历完成后设置合并
				if (j == (headerRows-1)) {
					verticalVlaue = null;
					
					if (verticalMerge!=null) {
						super.setMerge(sheet, verticalMerge.getFirstRow(), verticalMerge.getLastRow(), verticalMerge.getFirstCol(), verticalMerge.getLastCol());
						verticalMerge = null;
					}
				}
			}
		}
		
		// 最后一列遍历完成后设置合并
		if (verticalMerge!=null) {
			super.setMerge(sheet, verticalMerge.getFirstRow(), verticalMerge.getLastRow(), verticalMerge.getFirstCol(), verticalMerge.getLastCol());
			verticalVlaue = null;
			verticalMerge = null;
		}
	}
	
}
