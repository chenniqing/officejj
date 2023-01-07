package cn.javaex.officejj.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

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
	 * 默认值
	 *     填写在Excel上显示的内容
	 * @return
	 */
	public String defaultValue() default "";
	
}
