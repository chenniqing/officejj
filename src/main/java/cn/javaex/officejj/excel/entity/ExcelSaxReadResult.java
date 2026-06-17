package cn.javaex.officejj.excel.entity;

/**
 * Excel SAX读取结果。
 *
 * @author 陈霓清
 */
public class ExcelSaxReadResult {
	private int sheetCount;      // 已解析Sheet数量
	private int rowCount;        // 已回调数据行数量
	private boolean cancelled;   // 是否被取消

	public int getSheetCount() {
		return sheetCount;
	}

	public void setSheetCount(int sheetCount) {
		this.sheetCount = sheetCount;
	}

	public int getRowCount() {
		return rowCount;
	}

	public void setRowCount(int rowCount) {
		this.rowCount = rowCount;
	}

	public boolean isCancelled() {
		return cancelled;
	}

	public void setCancelled(boolean cancelled) {
		this.cancelled = cancelled;
	}
}
