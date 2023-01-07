package cn.javaex.officejj.common.util;

import java.util.Map;

/**
 * Map工具类
 * 
 * @author 陈霓清
 */
public class MapHandler {
	
	/**
	 * 获取map中第一个数据值
	 * @param <K>
	 * @param <V>
	 * @param map
	 * @return
	 */
	public static <K, V> V getFirstOrNull(Map<K, V> map) {
		V obj = null;
		
		for (Map.Entry<K, V> entry : map.entrySet()) {
			obj = entry.getValue();
			if (obj!=null) {
				break;
			}
		}
		
		return obj;
	}
	
}
