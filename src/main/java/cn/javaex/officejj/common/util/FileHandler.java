package cn.javaex.officejj.common.util;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.net.URLEncoder;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.util.IOUtils;

/**
 * 文件工具类
 * 
 * @author 陈霓清
 */
public class FileHandler {
	// 默认缓冲区大小
	private static final int BUFFER_SIZE = 2048;

	/**
	 * 文件下载（不重命名）
	 * @param filePath  文件的路径（带具体的文件名）
	 *                    如果是相对路径，则认为是项目的同级目录
	 *                      如果是springboot源码运行，则认为相对路径是项目名文件夹下的路径
	 */
	public static void downloadFile(HttpServletResponse response, String filePath) {
		downloadFile(response, filePath, null);
	}
	
	/**
	 * 文件下载
	 * @param filePath      文件的路径（带具体的文件名）
	 *                        如果是相对路径，则认为是项目的同级目录
	 *                          如果是springboot源码运行，则认为相对路径是项目名文件夹下的路径
	 * @param newFileName   重命名文件名称（带后缀）
	 */
	public static void downloadFile(HttpServletResponse response, String filePath, String newFileName) {
		// 传入的路径是否是绝对路径
		boolean isAbsolutePath = PathHandler.isAbsolutePath(filePath);
		// 存储文件的物理路径
		String fileAbsolutePath = "";
		if (isAbsolutePath) {
			fileAbsolutePath = filePath;
		} else {
			String projectPath = PathHandler.getProjectPath();
			fileAbsolutePath = projectPath + File.separator + filePath;
		}
		
		File file = new File(fileAbsolutePath);
		if (newFileName==null || newFileName.length()==0) {
			newFileName = file.getName();
		}
		
		BufferedInputStream bis = null;
		BufferedOutputStream bos = null;
		
		try {
			response.setContentType("application/octet-stream");
			response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(newFileName, "UTF-8"));
			response.setHeader("Content-Length", String.valueOf(file.length()));
			
			bis = new BufferedInputStream(new FileInputStream(file));
			bos = new BufferedOutputStream(response.getOutputStream());
			byte[] buff = new byte[BUFFER_SIZE];
			while (true) {
				int bytesRead;
				
				if (-1 == (bytesRead=bis.read(buff, 0, buff.length))) {
					break;
				}
				
				bos.write(buff, 0, bytesRead);
			}
			
			bos.flush();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(bos);
			IOUtils.closeQuietly(bis);
		}
	}
	
}
