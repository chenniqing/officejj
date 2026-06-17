package cn.javaex.officejj.excel.help;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import cn.javaex.officejj.excel.ExcelUtils;
import cn.javaex.officejj.excel.entity.ExcelImportErrorMarkSetting;

/**
 * Excel导入失败行标记。
 *
 * @author 陈霓清
 */
public class ExcelImportErrorMarkHelper {
	private static final int EXCEL_MAX_COLUMN_WIDTH = 255;

	/**
	 * 标记导入失败行。
	 * @param workbook
	 * @param rowErrorMap 失败行号和错误信息，行号从1开始
	 * @param setting 标记配置
	 * @return
	 */
	public Workbook markErrorRows(Workbook workbook, Map<Integer, List<String>> rowErrorMap, ExcelImportErrorMarkSetting setting) {
		if (workbook==null) {
			throw new IllegalArgumentException("Workbook不能为空");
		}
		if (rowErrorMap==null || rowErrorMap.isEmpty()) {
			return workbook;
		}
		if (setting==null) {
			setting = new ExcelImportErrorMarkSetting();
		}
		this.validateSetting(workbook, setting);

		Sheet sheet = workbook.getSheetAt(setting.getSheetNum() - 1);
		Row headerRow = this.getOrCreateRow(sheet, setting.getHeaderRowNum() - 1);
		int errorColIndex = this.getErrorColIndex(headerRow, setting);

		Cell headerCell = headerRow.getCell(errorColIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
		headerCell.setCellValue(setting.getErrorHeader());
		headerCell.setCellStyle(this.buildHeaderStyle(workbook, headerCell.getCellStyle(), setting));

		Map<Short, CellStyle> styleCache = new HashMap<Short, CellStyle>();
		for (Map.Entry<Integer, List<String>> entry : rowErrorMap.entrySet()) {
			Integer excelRowNum = entry.getKey();
			if (excelRowNum==null || excelRowNum<=0) {
				throw new IllegalArgumentException("错误行号必须从1开始：" + excelRowNum);
			}
			this.markRow(sheet, excelRowNum - 1, errorColIndex, entry.getValue(), styleCache, setting);
		}

		int width = Math.max(1, Math.min(EXCEL_MAX_COLUMN_WIDTH, setting.getErrorColumnWidth()));
		sheet.setColumnWidth(errorColIndex, width * 256);
		return workbook;
	}

	/**
	 * 校验配置。
	 * @param workbook
	 * @param setting
	 */
	private void validateSetting(Workbook workbook, ExcelImportErrorMarkSetting setting) {
		if (setting.getSheetNum()<=0 || setting.getSheetNum()>workbook.getNumberOfSheets()) {
			throw new IllegalArgumentException("Sheet序号不合法：" + setting.getSheetNum());
		}
		if (setting.getHeaderRowNum()<=0) {
			throw new IllegalArgumentException("表头行号必须从1开始：" + setting.getHeaderRowNum());
		}
		if (setting.getErrorColNum()<0) {
			throw new IllegalArgumentException("错误信息列号不能小于0：" + setting.getErrorColNum());
		}
		if (setting.getErrorHeader()==null || setting.getErrorHeader().trim().length()==0) {
			throw new IllegalArgumentException("错误信息列表头不能为空");
		}
	}

	/**
	 * 得到错误信息列索引。
	 * errorColNum指定时使用指定列；未指定时优先复用同名列，否则追加到表头最后。
	 * @param headerRow
	 * @param setting
	 * @return
	 */
	private int getErrorColIndex(Row headerRow, ExcelImportErrorMarkSetting setting) {
		if (setting.getErrorColNum()>0) {
			return setting.getErrorColNum() - 1;
		}

		short lastCellNum = headerRow.getLastCellNum();
		int endColIndex = Math.max(lastCellNum, 0);
		for (int i=0; i<endColIndex; i++) {
			String value = ExcelUtils.getCellValue(headerRow.getCell(i));
			if (setting.getErrorHeader().equals(value)) {
				return i;
			}
		}

		return endColIndex;
	}

	/**
	 * 标记单行错误。
	 * @param sheet
	 * @param rowIndex 行索引，从0开始
	 * @param errorColIndex 错误信息列索引，从0开始
	 * @param messageList 错误信息
	 * @param styleCache 样式缓存，避免重复创建样式
	 * @param setting 标记配置
	 */
	private void markRow(Sheet sheet, int rowIndex, int errorColIndex, List<String> messageList, Map<Short, CellStyle> styleCache, ExcelImportErrorMarkSetting setting) {
		Row row = this.getOrCreateRow(sheet, rowIndex);
		short lastCellNum = row.getLastCellNum();
		int endColIndex = Math.max(lastCellNum<0 ? 0 : lastCellNum, errorColIndex);
		for (int i=0; i<endColIndex; i++) {
			Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
			CellStyle sourceStyle = cell.getCellStyle();
			short styleIndex = sourceStyle==null ? (short) -1 : sourceStyle.getIndex();
			CellStyle errorRowStyle = styleCache.get(styleIndex);
			if (errorRowStyle==null) {
				errorRowStyle = cell.getSheet().getWorkbook().createCellStyle();
				if (sourceStyle!=null) {
					errorRowStyle.cloneStyleFrom(sourceStyle);
				}
				errorRowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				errorRowStyle.setFillForegroundColor(setting.getRowFillColor());
				styleCache.put(styleIndex, errorRowStyle);
			}
			cell.setCellStyle(errorRowStyle);
		}

		Cell errorCell = row.getCell(errorColIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
		errorCell.setCellValue(this.joinMessages(messageList));
		errorCell.setCellStyle(this.buildErrorMessageStyle(sheet.getWorkbook(), errorCell.getCellStyle(), setting));
	}

	/**
	 * 得到或创建行。
	 * @param sheet
	 * @param rowIndex
	 * @return
	 */
	private Row getOrCreateRow(Sheet sheet, int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row==null) {
			row = sheet.createRow(rowIndex);
		}
		return row;
	}

	/**
	 * 构建错误表头样式。
	 * @param workbook
	 * @param sourceStyle 原始样式
	 * @param setting 标记配置
	 * @return
	 */
	private CellStyle buildHeaderStyle(Workbook workbook, CellStyle sourceStyle, ExcelImportErrorMarkSetting setting) {
		CellStyle style = workbook.createCellStyle();
		if (sourceStyle!=null) {
			style.cloneStyleFrom(sourceStyle);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFillForegroundColor(setting.getHeaderFillColor());

		Font font = workbook.createFont();
		font.setBold(true);
		font.setColor(setting.getHeaderFontColor());
		style.setFont(font);

		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		this.setBorder(style);
		return style;
	}

	/**
	 * 构建错误信息单元格样式。
	 * @param workbook
	 * @param sourceStyle 原始样式
	 * @param setting 标记配置
	 * @return
	 */
	private CellStyle buildErrorMessageStyle(Workbook workbook, CellStyle sourceStyle, ExcelImportErrorMarkSetting setting) {
		CellStyle style = workbook.createCellStyle();
		if (sourceStyle!=null) {
			style.cloneStyleFrom(sourceStyle);
		}
		style.setWrapText(true);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFillForegroundColor(setting.getRowFillColor());
		this.setBorder(style);

		Font font = workbook.createFont();
		font.setColor(setting.getErrorFontColor());
		font.setBold(true);
		style.setFont(font);
		return style;
	}

	/**
	 * 设置细边框。
	 * @param style
	 */
	private void setBorder(CellStyle style) {
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
	}

	/**
	 * 拼接错误信息，过滤空消息，避免 String.join 遇到 null。
	 * @param messageList
	 * @return
	 */
	private String joinMessages(List<String> messageList) {
		if (messageList==null || messageList.isEmpty()) {
			return "";
		}

		StringBuilder sb = new StringBuilder();
		for (String message : messageList) {
			if (message==null || message.length()==0) {
				continue;
			}
			if (sb.length()>0) {
				sb.append("；");
			}
			sb.append(message);
		}
		return sb.toString();
	}
}
