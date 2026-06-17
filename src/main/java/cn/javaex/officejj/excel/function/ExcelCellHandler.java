package cn.javaex.officejj.excel.function;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Excel单元格处理器。
 * 可在写出后统一追加批注、超链接、样式、校验等自定义逻辑。
 *
 * @author 陈霓清
 */
@FunctionalInterface
public interface ExcelCellHandler {

	/**
	 * 处理单元格。
	 * @param sheet Sheet对象
	 * @param row 行对象
	 * @param cell 单元格对象
	 * @throws Exception
	 */
	void handle(Sheet sheet, Row row, Cell cell) throws Exception;
}
