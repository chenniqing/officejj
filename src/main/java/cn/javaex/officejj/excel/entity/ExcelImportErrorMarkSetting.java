package cn.javaex.officejj.excel.entity;

import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Excel导入错误标记配置。
 *
 * @author 陈霓清
 */
public class ExcelImportErrorMarkSetting {
	private int sheetNum = 1;                                  // 第几个Sheet，从1开始计算
	private int headerRowNum = 1;                              // 表头所在行，从1开始计算
	private int errorColNum = 0;                               // 错误信息列，从1开始计算；0表示自动追加或复用同名列
	private String errorHeader = "导入错误信息";                 // 错误信息列表头
	private int errorColumnWidth = 80;                         // 错误信息列宽，按Excel字符宽度计算
	private short rowFillColor = IndexedColors.ROSE.getIndex();
	private short headerFillColor = IndexedColors.RED.getIndex();
	private short headerFontColor = IndexedColors.WHITE.getIndex();
	private short errorFontColor = IndexedColors.RED.getIndex();

	/**
	 * 得到Sheet序号。
	 * @return
	 */
	public int getSheetNum() {
		return sheetNum;
	}

	/**
	 * 设置Sheet序号，从1开始计算。
	 * @param sheetNum
	 */
	public void setSheetNum(int sheetNum) {
		this.sheetNum = sheetNum;
	}

	/**
	 * 得到表头行号。
	 * @return
	 */
	public int getHeaderRowNum() {
		return headerRowNum;
	}

	/**
	 * 设置表头行号，从1开始计算。
	 * @param headerRowNum
	 */
	public void setHeaderRowNum(int headerRowNum) {
		this.headerRowNum = headerRowNum;
	}

	/**
	 * 得到错误信息列号。
	 * @return
	 */
	public int getErrorColNum() {
		return errorColNum;
	}

	/**
	 * 设置错误信息列号，从1开始计算；0表示自动追加或复用同名列。
	 * @param errorColNum
	 */
	public void setErrorColNum(int errorColNum) {
		this.errorColNum = errorColNum;
	}

	/**
	 * 得到错误信息列表头。
	 * @return
	 */
	public String getErrorHeader() {
		return errorHeader;
	}

	/**
	 * 设置错误信息列表头。
	 * @param errorHeader
	 */
	public void setErrorHeader(String errorHeader) {
		this.errorHeader = errorHeader;
	}

	/**
	 * 得到错误信息列宽。
	 * @return
	 */
	public int getErrorColumnWidth() {
		return errorColumnWidth;
	}

	/**
	 * 设置错误信息列宽，最大不超过Excel允许的255字符。
	 * @param errorColumnWidth
	 */
	public void setErrorColumnWidth(int errorColumnWidth) {
		this.errorColumnWidth = errorColumnWidth;
	}

	/**
	 * 得到错误行背景色。
	 * @return
	 */
	public short getRowFillColor() {
		return rowFillColor;
	}

	/**
	 * 设置错误行背景色。
	 * @param rowFillColor
	 */
	public void setRowFillColor(short rowFillColor) {
		this.rowFillColor = rowFillColor;
	}

	/**
	 * 得到错误表头背景色。
	 * @return
	 */
	public short getHeaderFillColor() {
		return headerFillColor;
	}

	/**
	 * 设置错误表头背景色。
	 * @param headerFillColor
	 */
	public void setHeaderFillColor(short headerFillColor) {
		this.headerFillColor = headerFillColor;
	}

	/**
	 * 得到错误表头字体颜色。
	 * @return
	 */
	public short getHeaderFontColor() {
		return headerFontColor;
	}

	/**
	 * 设置错误表头字体颜色。
	 * @param headerFontColor
	 */
	public void setHeaderFontColor(short headerFontColor) {
		this.headerFontColor = headerFontColor;
	}

	/**
	 * 得到错误信息字体颜色。
	 * @return
	 */
	public short getErrorFontColor() {
		return errorFontColor;
	}

	/**
	 * 设置错误信息字体颜色。
	 * @param errorFontColor
	 */
	public void setErrorFontColor(short errorFontColor) {
		this.errorFontColor = errorFontColor;
	}
}
