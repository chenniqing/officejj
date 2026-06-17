package cn.javaex.officejj.common.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;

/**
 * 图片处理工具类
 *
 * @author 陈霓清
 */
public class ImageHandler {

	/**
	 * 下载图片并转为流
	 * @param imageUrl
	 * @return
	 * @throws Exception
	 */
	public static InputStream downloadImageAsStream(String imageUrl) throws Exception {
		if (imageUrl==null || imageUrl.length()==0) {
			throw new IllegalArgumentException("图片地址不能为空");
		}

		String encodedUrl = imageUrl;
		if (!imageUrl.contains("?")) {
			// 将空格等特殊字符转换为URL编码，但保留http/https和主机部分
	        int pathIndex = imageUrl.indexOf("/", imageUrl.indexOf("//") + 2);
	        String domain = pathIndex<0 ? imageUrl : imageUrl.substring(0, pathIndex);
	        String path = pathIndex<0 ? "" : imageUrl.substring(pathIndex);
	        String encodedPath = path.length()==0 ? "" : encodeUrlPath(path);

	        encodedUrl = domain + encodedPath;
		}

        URL url = new URL(encodedUrl);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setConnectTimeout(5000);
        conn.setReadTimeout(5000);
        conn.connect();

        int responseCode = conn.getResponseCode();
        if (responseCode == HttpURLConnection.HTTP_OK) {
            return conn.getInputStream();
        } else {
            throw new RuntimeException("图片下载失败，HTTP响应码: " + responseCode);
        }
    }

    /**
     * 单独编码路径部分，避免整个URL都被编码。
     * @param path
     * @return
     * @throws Exception
     */
    private static String encodeUrlPath(String path) throws Exception {
        StringBuilder result = new StringBuilder();
        String[] segments = path.split("/");
        for (int i = 0; i < segments.length; i++) {
            if (segments[i].length() > 0) {
                result.append("/");
                result.append(URLEncoder.encode(segments[i], "UTF-8")
                        .replace("+", "%20")); // URLEncoder会把空格转成+，手动替换成%20
            }
        }
        // 保证根路径"/"正确
        if (path.endsWith("/")) {
            result.append("/");
        }
        return result.toString();
    }

	/**
	 * 获取图片流
	 *
	 * @param path
	 * @return
	 */
	public static InputStream getImageStream(String path) {
		try {
			if (path==null || path.length()==0) {
				throw new IllegalArgumentException("图片路径不能为空");
			}
			if (path.startsWith("http")) {
				return downloadImageAsStream(path);
			}
			else if (path.startsWith("resources:")) {
				path = path.replace("resources:", "");
				return PathHandler.getInputStreamFromResource(path);
			}
			else {
				// 存储文件的物理路径
				boolean isAbsolutePath = PathHandler.isAbsolutePath(path);
				String fileAbsolutePath = isAbsolutePath ? path : PathHandler.getProjectPath() + File.separator + path;
				return new FileInputStream(fileAbsolutePath);
			}
		} catch (Exception e) {
			throw new RuntimeException("读取图片失败：" + path, e);
		}
	}

}
