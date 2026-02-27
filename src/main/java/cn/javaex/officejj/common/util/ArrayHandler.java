package cn.javaex.officejj.common.util;

/**
 * 数组处理工具类
 * 
 * @author 陈霓清
 * @Date 2023年1月8日
 */
public class ArrayHandler {
	
	/**
	 * 数组中元素未找到的下标，值为-1
	 */
	public static final int INDEX_NOT_FOUND = -1;
	
	/**
	 * 数组中是否包含元素
	 *
	 * @param array 数组
	 * @param value 被检查的元素
	 * @return 是否包含
	 */
	public static boolean contains(int[] array, int value) {
		return indexOf(array, value) > INDEX_NOT_FOUND;
	}
	
	/**
	 * 返回数组中指定元素所在位置，未找到返回-1
	 *
	 * @param array 数组
	 * @param value 被检查的元素
	 */
	public static int indexOf(int[] array, int value) {
		if (null != array) {
			for (int i = 0; i < array.length; i++) {
				if (value == array[i]) {
					return i;
				}
			}
		}
		
		return INDEX_NOT_FOUND;
	}
	
}
