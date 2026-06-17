package cn.javaex.officejj.common.util;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

/**
 * 对象属性读取工具。
 * 支持 Map、JavaBean、普通字段、List/数组下标，主要用于模板占位符中的嵌套取值。
 *
 * @author 陈霓清
 */
public class PropertyHandler {

	/**
	 * 读取对象属性。
	 * 支持 user.name、user.dept.name、list[0].name、map.key 等写法。
	 * @param source 数据源，可以是 Map、JavaBean、List、数组
	 * @param path 属性路径
	 * @return
	 */
	public static Object getValue(Object source, String path) {
		if (source==null || path==null || path.trim().length()==0) {
			return null;
		}

		Object current = source;
		String[] parts = path.split("\\.");
		for (String part : parts) {
			if (current==null) {
				return null;
			}
			current = getSimpleValue(current, part);
		}
		return current;
	}

	/**
	 * 读取单段属性。
	 * 例如 name、items[0] 都属于单段属性。
	 * @param source 数据源
	 * @param part 单段属性
	 * @return
	 */
	private static Object getSimpleValue(Object source, String part) {
		if (part==null || part.length()==0) {
			return source;
		}

		int bracketIndex = part.indexOf('[');
		if (bracketIndex>=0 && part.endsWith("]")) {
			String name = part.substring(0, bracketIndex);
			String indexText = part.substring(bracketIndex + 1, part.length() - 1);
			Object value = name.length()==0 ? source : getSimpleValue(source, name);
			if (value==null) {
				return null;
			}
			int index = Integer.parseInt(indexText);
			if (value instanceof List) {
				List<?> list = (List<?>) value;
				return index>=0 && index<list.size() ? list.get(index) : null;
			}
			if (value.getClass().isArray()) {
				return index>=0 && index<java.lang.reflect.Array.getLength(value) ? java.lang.reflect.Array.get(value, index) : null;
			}
			return null;
		}

		if (source instanceof Map) {
			return ((Map<?, ?>) source).get(part);
		}

		Object value = invokeGetter(source, part);
		if (value!=null) {
			return value;
		}

		return readField(source, part);
	}

	/**
	 * 调用 JavaBean getter。
	 * @param source 数据源
	 * @param name 属性名
	 * @return
	 */
	private static Object invokeGetter(Object source, String name) {
		String suffix = Character.toUpperCase(name.charAt(0)) + name.substring(1);
		String[] methodNames = new String[] {"get" + suffix, "is" + suffix};
		for (String methodName : methodNames) {
			try {
				Method method = source.getClass().getMethod(methodName);
				method.setAccessible(true);
				return method.invoke(source);
			} catch (Exception e) {
				// 当前命名方式不存在时继续尝试其他读取方式。
			}
		}
		return null;
	}

	/**
	 * 读取字段值。
	 * @param source 数据源
	 * @param name 字段名
	 * @return
	 */
	private static Object readField(Object source, String name) {
		Class<?> clazz = source.getClass();
		while (clazz!=null && clazz!=Object.class) {
			try {
				Field field = clazz.getDeclaredField(name);
				field.setAccessible(true);
				return field.get(source);
			} catch (Exception e) {
				clazz = clazz.getSuperclass();
			}
		}
		return null;
	}

	/**
	 * 判断对象是否为逻辑真。
	 * 模板条件块使用该方法把常见类型转换成布尔含义。
	 * @param value 值
	 * @return
	 */
	public static boolean isTrue(Object value) {
		if (value==null) {
			return false;
		}
		if (value instanceof Boolean) {
			return (Boolean) value;
		}
		if (value instanceof Number) {
			return ((Number) value).doubleValue()!=0D;
		}
		String text = String.valueOf(value).trim();
		return text.length()>0 && !"false".equalsIgnoreCase(text) && !"0".equals(text) && !"否".equals(text);
	}

	/**
	 * 转成模板输出文本。
	 * @param value 值
	 * @return
	 */
	public static String toText(Object value) {
		return value==null ? "" : String.valueOf(value);
	}
}
