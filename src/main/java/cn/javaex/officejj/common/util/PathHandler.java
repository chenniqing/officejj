package cn.javaex.officejj.common.util;

import java.io.File;
import java.io.InputStream;

import javax.servlet.http.HttpServletRequest;

import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

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
	 * 获取项目所在磁盘的文件夹路径，并设置临时目录
	 * @return
	 */
	public static String getFolderPath() {
		String projectPath = getProjectPath();
		String folderPath = projectPath + File.separator + "temp_download";
		File file = new File(folderPath);
		file.mkdirs();
		
		return folderPath;
	}
	
	/**
	 * 获取项目所在磁盘的文件夹路径
	 * @return
	 */
	public static String getProjectPath() {
		try {
			HttpServletRequest request = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getRequest();
			// 获取地址内容，原路径（项目名）
			String realPath = request.getSession().getServletContext().getRealPath("/");
			// tomcat webapps下部署
			if (realPath!=null && realPath.length()>0 && realPath.contains("apache-tomcat")) {
				String path = request.getContextPath();    // 项目名
				path = path.replace("/", File.separator) + File.separator;
				return realPath.replace(path, "");
			}
			
			return System.getProperty("user.dir");
		} catch (Exception e) {
			return System.getProperty("user.dir");
		}
	}
	
	/**
	 * 获取服务路径
	 * @return
	 */
	public static String getServerPath() {
		HttpServletRequest request = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getRequest();
		
		String domain = request.getScheme() + "://" + request.getServerName();
		int port = request.getServerPort();
		
		return port==80 ? domain : domain + ":" + port;
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
