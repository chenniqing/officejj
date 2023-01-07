package cn.javaex.officejj.excel.help;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import cn.javaex.officejj.excel.entity.ExcelSetting;

/**
 * 设置类写入Excel
 * 
 * @author 陈霓清
 */
public class SheetSettingHelper extends SheetHelper {
	
	/**
	 * 创建内容
	 * @param sheet
	 * @param excelSetting
	 */
	@Override
	public void write(Sheet sheet, ExcelSetting excelSetting) {
		// 1.0 设置标题
		this.createTtile(sheet, excelSetting);
		
		// 2.0 设置表头
		this.createHeader(sheet, excelSetting);
		
		// 3.0 设置数据
		this.createData(sheet, excelSetting);
	}
	
	/**
	 * 设置标题
	 * @param sheet
	 * @param excelSetting
	 * @return
	 */
	private int createTtile(Sheet sheet, ExcelSetting excelSetting) {
		String title = excelSetting.getTitle();
		if (title==null || title.length()==0) {
			return 0;
		}
		
		Row row = sheet.createRow(0);
		// 行高
		int height = excelSetting.getTitleHeight();
		if (height>0) {
			row.setHeight((short) (height * BASE_ROW_HEIGHT));
		}
		
		// 标题样式
		CellStyle cellStyle = excelSetting.getCellStyle().createTitleStyle(sheet.getWorkbook());
		
		// 设置单元格
		Cell cell = row.createCell(0);
		cell.setCellValue(title);
		cell.setCellStyle(cellStyle);
		
		int length = 0;
		// 头部数据
		List<String[]> headerList = excelSetting.getHeaderList();
		if (headerList!=null && headerList.isEmpty()==false) {
			length = headerList.get(0).length;
		}
		
		// 设置合并
		// 四个参数分别是：起始行、终止行、起始列、终止列（从0开始计算）
		CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, length-1);
		sheet.addMergedRegion(cellRangeAddress);
		
		return 1;
	}
	
	/**
	 * 设置头部
	 * @param sheet
	 * @param excelSetting
	 * @return 
	 */
	private int createHeader(Sheet sheet, ExcelSetting excelSetting) {
		int rowIndex = 0;
		String title = excelSetting.getTitle();
		if (title!=null && title.length()>0) {
			rowIndex = 1;
		}
		
		// 头部样式
		CellStyle cellStyle = excelSetting.getCellStyle().createHeaderStyle(sheet.getWorkbook());
		// 头部数据
		List<String[]> headerList = excelSetting.getHeaderList();
		if (headerList!=null && headerList.isEmpty()==false) {
			for (int i=0; i<headerList.size(); i++) {
				// 创建行
				Row row = sheet.createRow(rowIndex);
				// 行高
				int height = excelSetting.getHeaderHeight();
				if (height>0) {
					row.setHeight((short) (height * BASE_ROW_HEIGHT));
				}
				
				String[] headerArr = headerList.get(i);
				for (int j=0; j<headerArr.length; j++) {
					// 设置单元格
					Cell cell = row.createCell(j);
					sheet.setColumnWidth(j, excelSetting.getColumnWidth() * BASE_COLUMN_WIDTH);
					cell.setCellValue(headerArr[j]);
					cell.setCellStyle(cellStyle);
				}
				
				rowIndex++;
			}
		}
		
		return rowIndex;
	}
	
	/**
	 * 设置数据
	 * @param sheet
	 * @param excelSetting
	 */
	private void createData(Sheet sheet, ExcelSetting excelSetting) {
		int dataRowIndex = 0;    // 数据行的起始索引
		
		String title = excelSetting.getTitle();
		if (title!=null && title.length()>0) {
			dataRowIndex += 1;
		}
		// 头部数据
		List<String[]> headerList = excelSetting.getHeaderList();
		if (headerList!=null && headerList.isEmpty()==false) {
			dataRowIndex += headerList.size();
		}
		
		// 数据样式
		CellStyle cellStyle = excelSetting.getCellStyle().createDataStyle(sheet.getWorkbook());
		// 数据
		List<String[]> dataList = excelSetting.getDataList();
		
		if (dataList!=null && dataList.isEmpty()==false) {
			int len = dataList.size();
			for (int i=0; i<len; i++) {
				// 创建行
				Row row = sheet.createRow(i + dataRowIndex);
				// 行高
				int height = excelSetting.getDataHeight();
				if (height>0) {
					row.setHeight((short) (height * BASE_ROW_HEIGHT));
				}
				
				// 得到每一行的数据
				String[] data = dataList.get(i);
				for (int j=0; j<data.length; j++) {
					// 设置单元格
					Cell cell = row.createCell(j);
					cell.setCellValue(data[j]);
					cell.setCellStyle(cellStyle);
				}
			}
		}
	}
}
