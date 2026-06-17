package cn.javaex.officejj.excel.help;

import java.util.ArrayList;
import java.util.List;

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
		this.setAutoMergeCol(sheet, colNum, firstRow, lastRow, clazz, new int[0]);
	}

	/**
	 * 自动合并列，并用指定依赖列作为合并边界。
	 * 例如第2列班主任按第1列班级拆分合并时，mergeByColNumArr 传入 0。
	 *
	 * @param sheet
	 * @param colNum      第几列（从0开始计算）
	 * @param firstRow    起始行（从0开始计算）
	 * @param lastRow     终止行（从0开始计算）
	 * @param clazz
	 * @param mergeByColNumArr 依赖列（从0开始计算），当前列值和依赖列值都相同才合并
	 */
	@Override
	public void setAutoMergeCol(Sheet sheet, int colNum, int firstRow, int lastRow, Class<?> clazz, int... mergeByColNumArr) {
		if (sheet==null || firstRow>=lastRow) {
			return;
		}

		int mergeRowBegin = firstRow;
		List<String> mergeKey = this.getMergeKey(sheet, firstRow, colNum, mergeByColNumArr);

		for (int i=firstRow+1; i<=lastRow; i++) {
			List<String> currentKey = this.getMergeKey(sheet, i, colNum, mergeByColNumArr);
			if (mergeKey.equals(currentKey)) {
				continue;
			}

			if (i-1>mergeRowBegin) {
				this.setMerge(sheet, mergeRowBegin, i-1, colNum, colNum, clazz);
			}
			mergeRowBegin = i;
			mergeKey = currentKey;
		}

		// 处理最后一组内容，只有跨越两行以上才需要真正合并。
		if (lastRow>mergeRowBegin) {
			this.setMerge(sheet, mergeRowBegin, lastRow, colNum, colNum, clazz);
		}
	}

	/**
	 * 生成合并判断键：当前列值 + 依赖列值。
	 * 使用列表比较可以避免简单字符串拼接导致的边界歧义。
	 * @param sheet
	 * @param rowIndex
	 * @param colNum
	 * @param mergeByColNumArr
	 * @return
	 */
	private List<String> getMergeKey(Sheet sheet, int rowIndex, int colNum, int... mergeByColNumArr) {
		List<String> list = new ArrayList<String>();
		list.add(ExcelUtils.readExcel(sheet, rowIndex + 1, colNum + 1));
		if (mergeByColNumArr!=null) {
			for (int mergeByColNum : mergeByColNumArr) {
				list.add(ExcelUtils.readExcel(sheet, rowIndex + 1, mergeByColNum + 1));
			}
		}

		return list;
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
