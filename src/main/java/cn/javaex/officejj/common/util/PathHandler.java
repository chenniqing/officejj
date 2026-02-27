package cn.javaex.officejj.common.util;

import java.io.InputStream;

/**
 * 路径工具类
 */
public class PathHandler {
	
	/**
	 * 路径转换
	 * @param str
	 * @return
	 */
	public static String slashify(String str) {
		return str.replace('\\', '/');
	}
	
	/**
	 * 判断是否是绝对路径
	 * @param path
	 * @return
	 */
	public static boolean isAbsolutePath(String path) {
		path = path.substring(0, 2);
		return path.startsWith("/") || path.endsWith(":");
	}

	/**
	 * 获取项目所在磁盘的文件夹路径
	 * @return
	 */
	public static String getProjectPath() {
		return System.getProperty("user.dir");
	}

	/**
	 * 根据resources下的文件路径获取流
	 * @param path    例如：template/excel/模板.xlsx
	 * @return
	 */
	public static InputStream getInputStreamFromResource(String path) {
		if (path.startsWith("/")) {
			path = path.substring(1);
		}
		
		return Thread.currentThread().getContextClassLoader().getResourceAsStream(path);
	}
}
