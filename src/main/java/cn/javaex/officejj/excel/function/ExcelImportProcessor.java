package cn.javaex.officejj.excel.function;

import org.apache.poi.ss.usermodel.Workbook;

import cn.javaex.officejj.excel.entity.ExcelImportResult;

/**
 * Excel导入处理器。
 *
 * @param <T> 导入后需要返回给业务方的数据类型
 * @author 陈霓清
 */
@FunctionalInterface
public interface ExcelImportProcessor<T> {

	/**
	 * 执行业务导入。
	 * 调用方可在这里读取Excel、校验数据、写入数据库，并把失败行放入 ExcelImportResult.rowErrorMap。
	 * @param workbook 导入的Workbook，生命周期由officejj管理
	 * @return
	 * @throws Exception
	 */
	ExcelImportResult<T> process(Workbook workbook) throws Exception;
}
