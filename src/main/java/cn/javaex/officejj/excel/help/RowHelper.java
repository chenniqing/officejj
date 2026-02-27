package cn.javaex.officejj.excel.help;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import cn.javaex.officejj.excel.ExcelUtils;

/**
 * 行
 * 
 * @author 陈霓清
 */
public class RowHelper {

	/**
	 * 复制一行所有单元格（含所有样式）
	 * @param workbook
	 * @param src
	 * @param tgt
	 */
	public void copyRow(Workbook workbook, Row src, Row tgt) {
		if (src == null) {
			return;
		}
		tgt.setHeight(src.getHeight());
		for (int j = 0; j < src.getLastCellNum(); j++) {
			Cell srcCell = src.getCell(j);
			Cell tgtCell = tgt.createCell(j);
			if (srcCell == null) {
				continue;
			}
			
			switch (srcCell.getCellType()) {
				case STRING:
					tgtCell.setCellValue(srcCell.getStringCellValue());
					break;
				case NUMERIC:
					tgtCell.setCellValue(srcCell.getNumericCellValue());
					break;
				case BOOLEAN:
					tgtCell.setCellValue(srcCell.getBooleanCellValue());
					break;
				case FORMULA:
					tgtCell.setCellFormula(srcCell.getCellFormula());
					break;
				case ERROR:
					tgtCell.setCellErrorValue(srcCell.getErrorCellValue());
					break;
				default:
					tgtCell.setBlank();
			}
			CellStyle newStyle = workbook.createCellStyle();
			newStyle.cloneStyleFrom(srcCell.getCellStyle());
			tgtCell.setCellStyle(newStyle);
		}
	}

	/**
	 * 复制行及其数据
	 * 
	 * @param sheet
	 * @param oldRow
	 * @param newRow
	 */
	public void copyRow(Sheet sheet, Row oldRow, Row newRow) {
		// 复制行高
		if (oldRow.getHeight()>=0) {
			newRow.setHeight(oldRow.getHeight());
		}
		
		// 循环复制单元格
		for (int i=oldRow.getFirstCellNum(); i<oldRow.getLastCellNum(); i++) {
			Cell oldCell = oldRow.getCell(i);
			Cell newCell = newRow.getCell(i);
			
			if (oldCell!=null) {
				if (newCell==null) {
					newCell = newRow.createCell(i);
				}
				
				// 复制单元格和样式
				this.copyCell(oldCell, newCell);
			}
		}
		
		// 复制合并单元格的样式
		for (int i=0; i<sheet.getNumMergedRegions(); i++) {
			CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
			if (cellRangeAddress.getFirstRow()==oldRow.getRowNum()) {
				CellRangeAddress newCellRangeAddress =
						new CellRangeAddress(
							newRow.getRowNum(),
							(newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
							cellRangeAddress.getFirstColumn(),
							cellRangeAddress.getLastColumn()
						);
				sheet.addMergedRegionUnsafe(newCellRangeAddress);
			}
		}
	}

	/**
	 * 复制单元格和样式
	 * @param oldCell
	 * @param newCell
	 */
	private void copyCell(Cell oldCell, Cell newCell) {
		// 复制样式
		newCell.setCellStyle(oldCell.getCellStyle());
		
		// 复制值
		newCell.setCellValue(ExcelUtils.getCellValue(oldCell));
	}

}
