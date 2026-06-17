package cn.javaex.officejj.excel.function;

/**
 * Excel读取进度监听器。
 *
 * @author 陈霓清
 */
@FunctionalInterface
public interface ExcelReadProgressListener {

	/**
	 * 读取进度回调。
	 * @param sheetNum 当前Sheet序号，从1开始
	 * @param rowCount 已读取数据行数
	 */
	void onProgress(int sheetNum, int rowCount);
}
