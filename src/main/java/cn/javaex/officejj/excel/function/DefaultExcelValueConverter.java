package cn.javaex.officejj.excel.function;

import java.lang.reflect.Field;

import cn.javaex.officejj.excel.annotation.ExcelCell;

/**
 * 默认字段转换器。
 * 导入时不改变原始文本，实际类型转换由 officejj 内置逻辑继续处理；导出时继承接口默认实现，保持字段原值写出。
 *
 * @author 陈霓清
 */
public class DefaultExcelValueConverter implements ExcelValueConverter {

	@Override
	public Object convert(String cellValue, Field field, ExcelCell excelCell) {
		return cellValue;
	}
}
