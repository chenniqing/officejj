package cn.javaex.officejj.common.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;

/**
 * 图片处理工具类
 * 
 * @author 陈霓清
 */
public class ImageHandler {
	
	/**
	 * 获取图片流
	 * 
	 * @param path
	 * @return
	 */
	public static InputStream getImageStream(String path) {
		try {
			if (path.startsWith("http")) {
				HttpURLConnection connection = (HttpURLConnection) new URL(path).openConnection();
				connection.setReadTimeout(1000);
				connection.setConnectTimeout(1000);
				connection.setRequestMethod("GET");
				if (connection.getResponseCode() == HttpURLConnection.HTTP_OK) {
					return connection.getInputStream();
				}
			}
			else if (path.startsWith("resources:")) {
				path = path.replace("resources:", "");
				return PathHandler.getInputStreamFromResource(path);
			}
			else {
				boolean isAbsolutePath = PathHandler.isAbsolutePath(path);
				
				// 存储文件的物理路径
				String fileAbsolutePath = "";
				if (isAbsolutePath) {
					fileAbsolutePath = path;
				} else {
					String projectPath = PathHandler.getProjectPath();
					fileAbsolutePath = projectPath + File.separator + path;
				}
				
				return new FileInputStream(fileAbsolutePath);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return null;
	}
	
}
