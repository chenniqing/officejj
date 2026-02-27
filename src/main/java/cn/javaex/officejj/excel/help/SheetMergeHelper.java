package cn.javaex.officejj.excel.help;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import cn.javaex.officejj.excel.ExcelUtils;
import cn.javaex.officejj.excel.entity.TransversalMerge;
import cn.javaex.officejj.excel.entity.VerticalMerge;
import cn.javaex.officejj.excel.style.ICellStyle;

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
	 * 返回：如果cell在某合并区则返回对应region，否则null
	 * @param sheet
	 * @param rowIdx
	 * @param colIdx
	 * @return
	 */
	public CellRangeAddress getMergedRegion(Sheet sheet, int rowIdx, int colIdx) {
		for (CellRangeAddress range : sheet.getMergedRegions()) {
			if (range.isInRange(rowIdx, colIdx)) {
				return range;
			}
		}
		return null;
	}

	/**
	 * 设置表头合并
	 * @param sheet
	 * @param rowIndex
	 * @param headerRows
	 * @throws Exception 
	 */
	public void setHeaderMerge(Sheet sheet, int rowIndex, int headerRows) throws Exception {
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
	 * @throws Exception 
	 */
	private void setTransverseMerge(Sheet sheet, int rowIndex, int headerRows) throws Exception {
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
							this.setMerge(sheet, transversalMerge.getFirstRow(), transversalMerge.getLastRow(), transversalMerge.getFirstCol(), transversalMerge.getLastCol(), null);
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
						this.setMerge(sheet, transversalMerge.getFirstRow(), transversalMerge.getLastRow(), transversalMerge.getFirstCol(), transversalMerge.getLastCol(), null);
						transverseVlaue = null;
						transversalMerge = null;
					}
				}
			}
		}
		
		// 最后一行遍历完成后设置横向合并
		if (transversalMerge!=null) {
			this.setMerge(sheet, transversalMerge.getFirstRow(), transversalMerge.getLastRow(), transversalMerge.getFirstCol(), transversalMerge.getLastCol(), null);
			transverseVlaue = null;
			transversalMerge = null;
		}
	}
	
	/**
	 * 设置纵向合并
	 * @param sheet
	 * @param rowIndex
	 * @param headerRows
	 * @throws Exception 
	 */
	private void setVerticalMerge(Sheet sheet, int rowIndex, int headerRows) throws Exception {
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
							this.setMerge(sheet, verticalMerge.getFirstRow(), verticalMerge.getLastRow(), verticalMerge.getFirstCol(), verticalMerge.getLastCol(), null);
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
						this.setMerge(sheet, verticalMerge.getFirstRow(), verticalMerge.getLastRow(), verticalMerge.getFirstCol(), verticalMerge.getLastCol(), null);
						verticalMerge = null;
					}
				}
			}
		}
		
		// 最后一列遍历完成后设置合并
		if (verticalMerge!=null) {
			this.setMerge(sheet, verticalMerge.getFirstRow(), verticalMerge.getLastRow(), verticalMerge.getFirstCol(), verticalMerge.getLastCol(), null);
			verticalVlaue = null;
			verticalMerge = null;
		}
	}
	
	/**
	 * 自动合并列
	 * 
	 * @param sheet
	 * @param colNum      第几列（从0开始计算）
	 * @param firstRow    起始行（从0开始计算）
	 * @param lastRow     终止行（从0开始计算）
	 * @param clazz
	 */
	@Override
	public void setAutoMergeCol(Sheet sheet, int colNum, int firstRow, int lastRow, Class<?> clazz) {
		// 合并定义
		int mergeRowBegin = 0;
		int mergeRowEnd = 0;
		String mergeName = null;    // 合并内容
		
		for (int i = firstRow; i <= lastRow; i++) {
			String content = ExcelUtils.readExcel(sheet, i + 1, colNum + 1);
			
			if (mergeName != null && content.equals(mergeName)) {
				mergeRowEnd++;
			} else {
				if (mergeRowEnd > mergeRowBegin) {
					this.setMerge(sheet, mergeRowBegin-1, mergeRowEnd-1, colNum, colNum, clazz);
				}
				mergeRowBegin = i + 1;
				mergeRowEnd = mergeRowBegin;
				mergeName = content;
			}
		}
		
		// 处理最后2行相等的情况
		if (mergeRowEnd > mergeRowBegin) {
			this.setMerge(sheet, mergeRowBegin-1, mergeRowEnd-1, colNum, colNum, clazz);
		}
	}
	
	/**
	 * 自动合并行
	 * @param sheet
	 * @param rowNum      第几行（从0开始计算）
	 * @param firstCol    起始列（从0开始计算）
	 * @param lastCol     终止列（从0开始计算）
	 * @param clazz
	 */
	@Override
	public void setAutoMergeRow(Sheet sheet, int rowNum, int firstCol, Integer lastCol, Class<?> clazz) {
		int mergeColBegin = 0;
		int mergeColEnd = 0;
		String mergeName = null; // 合并内容
		
		for (int j = firstCol; j <= lastCol; j++) {
			String content = ExcelUtils.readExcel(sheet, rowNum + 1, j + 1);
			
			if (mergeName != null && content.equals(mergeName)) {
				mergeColEnd++;
			} else {
				if (mergeColEnd > mergeColBegin) {
					this.setMerge(sheet, rowNum, rowNum, mergeColBegin, mergeColEnd, clazz);
				}
				mergeColBegin = j;
				mergeColEnd = j;
				mergeName = content;
			}
		}
		
		// 处理最后一组内容
		if (mergeColEnd > mergeColBegin) {
			this.setMerge(sheet, rowNum, rowNum, mergeColBegin, mergeColEnd, clazz);
		}
	}
	
	/**
	 * 设置合并
	 * @param sheet
	 * @param firstRow    起始行（从0开始计算）
	 * @param lastRow     终止行（从0开始计算）
	 * @param firstCol    起始列（从0开始计算）
	 * @param lastCol     终止列（从0开始计算）
	 * @param clazz
	 */
	@Override
	public void setMerge(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, Class<?> clazz) {
		CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		sheet.addMergedRegion(cellRangeAddress);
		
		// 合并样式
		if (clazz != null) {
			try {
				ICellStyle styleProvider = (ICellStyle) clazz.getDeclaredConstructor().newInstance();
				
				Row row = sheet.getRow(firstRow);
				if (row != null) {
					Cell cell = row.getCell(firstCol);
					if (cell != null) {
						cell.setCellStyle(styleProvider.createDataStyle(sheet.getWorkbook()));
					}
				}
			} catch (Exception e) {
				throw new RuntimeException("设置单元格合并样式失败", e);
			}
		}
	}
	
}
