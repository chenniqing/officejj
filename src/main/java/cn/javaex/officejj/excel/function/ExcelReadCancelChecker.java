package cn.javaex.officejj.excel.function;

/**
 * Excel读取取消检查器。
 * 大文件导入时可用该接口接入任务中心的取消状态。
 *
 * @author 陈霓清
 */
@FunctionalInterface
public interface ExcelReadCancelChecker {

	/**
	 * 是否取消读取。
	 * @return
	 */
	boolean isCancelled();
}
