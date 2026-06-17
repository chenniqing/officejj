package cn.javaex.officejj.excel.function;

import java.lang.reflect.Field;

import cn.javaex.officejj.excel.annotation.ExcelCell;

/**
 * Excel字段转换器。
 * 用于处理字典、枚举、复杂格式等业务值转换；导入失败时抛出异常即可进入导入错误收集。
 *
 * @author 陈霓清
 */
public interface ExcelValueConverter {

	/**
	 * 导入时把单元格文本转换成字段值。
	 * @param cellValue 单元格文本
	 * @param field 当前字段
	 * @param excelCell 字段上的ExcelCell注解，可能为空
	 * @return
	 * @throws Exception
	 */
	Object convert(String cellValue, Field field, ExcelCell excelCell) throws Exception;

	/**
	 * 导出时把字段原始值转换成最终写入单元格的值。
	 * 默认直接返回原值，兼容只实现导入转换的旧业务转换器；需要导出自定义展示时覆盖本方法即可。
	 * @param fieldValue 字段原始值，可能为空
	 * @param field 当前字段
	 * @param excelCell 字段上的ExcelCell注解，可能为空
	 * @return 转换后写入Excel的值
	 * @throws Exception
	 */
	default Object convertToExcel(Object fieldValue, Field field, ExcelCell excelCell) throws Exception {
		return fieldValue;
	}
}
