package cn.javaex.officejj.excel.entity;

/**
 * Excel SAX读取配置。
 * SAX模式只解析xlsx文件，适合大文件导入时按批次处理，避免一次性把全部数据放进内存。
 *
 * @author 陈霓清
 */
public class ExcelSaxReadSetting {
	private int sheetNum = 1;               // 第几个Sheet，从1开始；0表示读取全部Sheet
	private int startRowNum = 1;            // 从第几行开始读取，从1开始
	private int batchSize = 500;            // 每批回调多少行
	private int maxRows = 0;                // 最大读取行数，0表示不限制
	private boolean readEmptyRow = false;   // 是否回调空行

	public int getSheetNum() {
		return sheetNum;
	}

	public void setSheetNum(int sheetNum) {
		this.sheetNum = sheetNum;
	}

	public int getStartRowNum() {
		return startRowNum;
	}

	public void setStartRowNum(int startRowNum) {
		this.startRowNum = startRowNum;
	}

	public int getBatchSize() {
		return batchSize;
	}

	public void setBatchSize(int batchSize) {
		this.batchSize = batchSize;
	}

	public int getMaxRows() {
		return maxRows;
	}

	public void setMaxRows(int maxRows) {
		this.maxRows = maxRows;
	}

	public boolean isReadEmptyRow() {
		return readEmptyRow;
	}

	public void setReadEmptyRow(boolean readEmptyRow) {
		this.readEmptyRow = readEmptyRow;
	}
}
