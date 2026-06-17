package cn.javaex.officejj.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import cn.javaex.officejj.excel.function.DefaultExcelValueConverter;
import cn.javaex.officejj.excel.function.ExcelValueConverter;

/**
 * Excel单元格
 * 
 * @author 陈霓清
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCell {
	
	/**
	 * 表头（对应数据库字段，即该列的名称）
	 *     多行表头时自动合并
	 *     例如：{"表头1", "表头2"}  或  "表头"
	 * @return
	 */
	public String[] name() default {};
	
	/**
	 * 值替换
	 *     replace={"1_男", "0_女"}：表示数据库值为“1”时，替换为“男”，值为“0”时，替换为“女”
	 * @return
	 */
	public String[] replace() default {};
	
	/**
	 * 排序，从 1 开始计算
	 *     如果都缺省的话（即默认值0），则按照成员变量的顺序自动排序
	 * @return
	 */
	public int sort() default 0;
	
	/**
	 * 导出时，每列的宽度
	 *     单位为字符。1个汉字=2个字符
	 * @return
	 */
	public int width() default 16;
	
	/**
	 * 格式化
	 *     例如：format="yyyy-MM-dd"
	 * @return
	 */
	public String format() default "";
	
	/**
	 * 类型，默认都是文本
	 *     例如：type="image"    表示该列是图片列
	 * @return
	 */
	public String type() default "";
	
	/**
	 * 多少列合并成一个组
	 *     超过1时有效，自动向后合并
	 * @return
	 */
	public int group() default 1;
	
	/**
	 * 合并成一个组时的分隔符
	 * @return
	 */
	public String separator() default " / ";

	/**
	 * 是否纵向自动合并相邻相同值。
	 *     适用于导出明细列表时，把连续相同的班级、部门等父级字段自动合并成一个单元格。
	 * @return
	 */
	public boolean mergeRow() default false;

	/**
	 * 纵向合并的依赖列（从1开始计算）。
	 *     当前列只有在自身值相同，并且依赖列值也相同时才会继续合并。
	 *     例如班主任列 mergeBy={1}，表示同一个班主任跨多个班级时，会按第1列班级边界拆开合并。
	 * @return
	 */
	public int[] mergeBy() default {};

	/**
	 * 导入/导出字段转换器。
	 *     适用于字典、枚举、复杂日期、业务编码等内置类型转换无法覆盖的场景。
	 *     导入时实现 convert，导出时实现 convertToExcel；未覆盖导出方法时保持原字段值写出。
	 *     转换器抛出异常时，导入流程会收集为当前行错误，导出流程会直接向上抛出便于调用方定位。
	 * @return
	 */
	public Class<? extends ExcelValueConverter> converter() default DefaultExcelValueConverter.class;
	
	/**
	 * 默认值
	 *     填写在Excel上显示的内容
	 * @return
	 */
	public String defaultValue() default "";
	
}
