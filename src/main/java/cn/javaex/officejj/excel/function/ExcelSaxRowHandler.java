package cn.javaex.officejj.excel.function;

import java.util.List;

import cn.javaex.officejj.excel.entity.ExcelSaxRow;

/**
 * Excel SAX批次行处理器。
 *
 * @author 陈霓清
 */
@FunctionalInterface
public interface ExcelSaxRowHandler {

	/**
	 * 处理一批Excel行。
	 * @param rowList 行数据
	 * @throws Exception
	 */
	void handle(List<ExcelSaxRow> rowList) throws Exception;
}
