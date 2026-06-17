package cn.javaex.officejj.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel样式
 * 
 * @author 陈霓清
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelStyle {
	
	/**
	 * 自定义样式实现类名
	 * @return
	 */
	public String cellStyle() default "cn.javaex.officejj.excel.style.DefaultCellStyle";
	
	/**
	 * 标题栏高度
	 * @return
	 */
	public int titleHeight() default 0;
	
	/**
	 * 表头高度
	 * @return
	 */
	public int headerHeight() default 0;
	
	/**
	 * 数据行高度
	 * @return
	 */
	public int dataHeight() default 0;

	/**
	 * 是否根据数据行文本内容自动调整行高。
	 *     dataHeight 大于0时认为用户指定了固定行高，此配置不生效。
	 *     注解导出默认开启，避免长文本在导出结果中被行高遮挡。
	 *     如果业务需要固定版式，可显式设置 autoDataHeight=false 或设置 dataHeight。
	 * @return
	 */
	public boolean autoDataHeight() default true;

	/**
	 * 自动行高的最大高度，0表示不限制。
	 *     用于避免极长文本把单行撑得过高。
	 * @return
	 */
	public int maxDataHeight() default 0;

	/**
	 * 是否根据表头和数据内容自动调整列宽。
	 *     默认关闭，保持 @ExcelCell(width) 的固定列宽行为不变。
	 *     常用于 autoDataHeight=false 的列表导出：不自动撑高行高，但让列宽在可控范围内适配内容。
	 * @return
	 */
	public boolean autoColumnWidth() default false;

	/**
	 * 自动列宽的最小宽度，单位为字符宽度。
	 *     autoColumnWidth=true 时生效，0 表示使用 Excel 允许的最小安全宽度。
	 * @return
	 */
	public int minColumnWidth() default 0;

	/**
	 * 自动列宽的最大宽度，单位为字符宽度。
	 *     autoColumnWidth=true 时生效，0 表示不额外限制，最多不超过 Excel 允许的最大列宽。
	 * @return
	 */
	public int maxColumnWidth() default 0;
	
}
