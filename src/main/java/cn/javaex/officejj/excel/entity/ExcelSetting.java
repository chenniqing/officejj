package cn.javaex.officejj.excel.entity;

import java.util.List;

import cn.javaex.officejj.excel.style.DefaultCellStyle;
import cn.javaex.officejj.excel.style.ICellStyle;

/**
 * Excel配置类
 * 
 * @author 陈霓清
 */
public class ExcelSetting {
	private String sheetName;                                 // sheet页名称
	private String title;                                     // 顶部标题/说明
	private List<String[]> headerList;                        // 表头
	private List<String[]> dataList;                          // 数据
	private int columnWidth = 16;                             // 列宽
	private int titleHeight = 0;                              // 标题栏高度
	private int headerHeight = 0;                             // 表头高度
	private int dataHeight = 0;                               // 数据行高度
	private ICellStyle cellStyle = new DefaultCellStyle();    // 单元格样式

	/**
	 * 得到Sheet页名称
	 * @return
	 */
	public String getSheetName() {
		return sheetName;
	}
	
	/**
	 * 设置Sheet页名称
	 * @param sheet1Name
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
	/**
	 * 得到顶部标题/说明
	 * @return
	 */
	public String getTitle() {
		return title;
	}
	/**
	 * 设置顶部标题/说明
	 * @param title
	 */
	public void setTitle(String title) {
		this.title = title;
	}
	/**
	 * 得到表头
	 * @return
	 */
	public List<String[]> getHeaderList() {
		return headerList;
	}
	/**
	 * 设置表头
	 * @param headerList
	 */
	public void setHeaderList(List<String[]> headerList) {
		this.headerList = headerList;
	}

	/**
	 * 得到数据
	 * @return
	 */
	public List<String[]> getDataList() {
		return dataList;
	}
	/**
	 * 设置数据
	 * @param demoList
	 */
	public void setDataList(List<String[]> dataList) {
		this.dataList = dataList;
	}
	
	/**
	 * 得到列宽
	 * @return
	 */
	public int getColumnWidth() {
		return columnWidth;
	}
	/**
	 * 设置列宽
	 * @param columnWidth
	 */
	public void setColumnWidth(int columnWidth) {
		this.columnWidth = columnWidth;
	}

	/**
	 * 得到标题栏高度
	 * @return
	 */
	public int getTitleHeight() {
		return titleHeight;
	}
	/**
	 * 设置标题栏高度
	 * @param titleHeight
	 */
	public void setTitleHeight(int titleHeight) {
		this.titleHeight = titleHeight;
	}

	/**
	 * 得到表头高度
	 * @return
	 */
	public int getHeaderHeight() {
		return headerHeight;
	}
	/**
	 * 设置表头高度
	 * @param headerHeight
	 */
	public void setHeaderHeight(int headerHeight) {
		this.headerHeight = headerHeight;
	}

	/**
	 * 得到数据行高度
	 * @return
	 */
	public int getDataHeight() {
		return dataHeight;
	}
	/**
	 * 设置数据行高度
	 * @param dataHeight
	 */
	public void setDataHeight(int dataHeight) {
		this.dataHeight = dataHeight;
	}
	
	/**
	 * 得到单元格样式
	 * @return
	 */
	public ICellStyle getCellStyle() {
		return cellStyle;
	}
	/**
	 * 设置单元格样式
	 * @param cellStyle
	 */
	public void setCellStyle(ICellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}
}
