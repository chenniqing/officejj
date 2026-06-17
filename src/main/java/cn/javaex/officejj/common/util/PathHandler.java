package cn.javaex.officejj.common.util;

import java.io.File;
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
		if (str==null) {
			return null;
		}
		return str.replace('\\', '/');
	}

	/**
	 * 判断是否是绝对路径
	 * @param path
	 * @return
	 */
	public static boolean isAbsolutePath(String path) {
		if (path==null || path.length()==0) {
			return false;
		}

		// File#isAbsolute 同时兼容 Windows 盘符路径、UNC 路径和类 Unix 绝对路径。
		return new File(path).isAbsolute();
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
		if (path==null || path.length()==0) {
			throw new IllegalArgumentException("资源路径不能为空");
		}
		if (path.startsWith("/")) {
			path = path.substring(1);
		}

		InputStream in = Thread.currentThread().getContextClassLoader().getResourceAsStream(path);
		if (in==null) {
			throw new IllegalArgumentException("资源文件不存在：" + path);
		}

		return in;
	}
}
