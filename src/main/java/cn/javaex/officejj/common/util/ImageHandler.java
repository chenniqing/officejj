package cn.javaex.officejj.common.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;

/**
 * 图片处理工具类
 *
 * @author 陈霓清
 */
public class ImageHandler {

	private static final String USER_INFO_ALLOWED_CHARS = "-._~!$&'()*+,;=:";
	private static final String PATH_ALLOWED_CHARS = "-._~!$&'()*+,;=:@/";
	private static final String QUERY_ALLOWED_CHARS = "-._~!$&'()*+,;=:@/?";

	/**
	 * 下载图片并转为流。
	 * @param imageUrl
	 * @return
	 * @throws Exception
	 */
	public static InputStream downloadImageAsStream(String imageUrl) throws Exception {
		if (imageUrl==null || imageUrl.length()==0) {
			throw new IllegalArgumentException("图片地址不能为空");
		}

		String encodedUrl = encodeHttpUrl(imageUrl);
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
	 * 编码 HTTP URL 中需要转义的字符，同时保留已经编码过的 %XX。
	 * 这样既能读取中文图片名，也不会把调用方传入的 %20 二次编码成 %2520。
	 * @param imageUrl
	 * @return
	 * @throws Exception
	 */
	private static String encodeHttpUrl(String imageUrl) throws Exception {
		URL url = new URL(imageUrl);
		StringBuilder encodedUrl = new StringBuilder();
		encodedUrl.append(url.getProtocol()).append("://");
		if (url.getUserInfo()!=null && url.getUserInfo().length()>0) {
			encodedUrl.append(encodeUrlPart(url.getUserInfo(), USER_INFO_ALLOWED_CHARS)).append("@");
		}
		encodedUrl.append(url.getHost());
		if (url.getPort()>=0) {
			encodedUrl.append(":").append(url.getPort());
		}
		encodedUrl.append(encodeUrlPart(url.getPath(), PATH_ALLOWED_CHARS));
		if (url.getQuery()!=null) {
			encodedUrl.append("?").append(encodeUrlPart(url.getQuery(), QUERY_ALLOWED_CHARS));
		}
		if (url.getRef()!=null) {
			encodedUrl.append("#").append(encodeUrlPart(url.getRef(), QUERY_ALLOWED_CHARS));
		}

		return encodedUrl.toString();
	}

	/**
	 * 按 URL 组件逐字符编码：ASCII 安全字符原样保留，中文、空格等字符按 UTF-8 转成百分号编码。
	 * 对已经存在的百分号编码直接放行，避免重复编码导致远端路径不匹配。
	 * @param value
	 * @param allowedChars
	 * @return
	 */
	private static String encodeUrlPart(String value, String allowedChars) {
		StringBuilder result = new StringBuilder();
		for (int i=0; i<value.length();) {
			int codePoint = value.codePointAt(i);
			if (codePoint=='%' && isEncodedPercent(value, i)) {
				result.append(value, i, i + 3);
				i += 3;
				continue;
			}
			if (isAllowedUrlChar(codePoint, allowedChars)) {
				result.append((char) codePoint);
			} else {
				byte[] bytes = new String(Character.toChars(codePoint)).getBytes(StandardCharsets.UTF_8);
				for (byte b : bytes) {
					result.append('%');
					String hex = Integer.toHexString(b & 0xFF).toUpperCase();
					if (hex.length()==1) {
						result.append('0');
					}
					result.append(hex);
				}
			}
			i += Character.charCount(codePoint);
		}

		return result.toString();
	}

	/**
	 * 判断当前位置是否已经是合法的百分号编码。
	 * @param value
	 * @param index
	 * @return
	 */
	private static boolean isEncodedPercent(String value, int index) {
		return index + 2<value.length()
				&& isHexChar(value.charAt(index + 1))
				&& isHexChar(value.charAt(index + 2));
	}

	/**
	 * 判断字符是否属于 URL 组件中允许原样保留的 ASCII 字符。
	 * @param codePoint
	 * @param allowedChars
	 * @return
	 */
	private static boolean isAllowedUrlChar(int codePoint, String allowedChars) {
		return codePoint>='a' && codePoint<='z'
				|| codePoint>='A' && codePoint<='Z'
				|| codePoint>='0' && codePoint<='9'
				|| codePoint<128 && allowedChars.indexOf((char) codePoint)>=0;
	}

	/**
	 * 判断字符是否为十六进制字符。
	 * @param c
	 * @return
	 */
	private static boolean isHexChar(char c) {
		return c>='0' && c<='9'
				|| c>='a' && c<='f'
				|| c>='A' && c<='F';
	}

	/**
	 * 判断是否为 HTTP/HTTPS 图片地址。
	 * @param path
	 * @return
	 */
	private static boolean isHttpUrl(String path) {
		return path.regionMatches(true, 0, "http://", 0, "http://".length())
				|| path.regionMatches(true, 0, "https://", 0, "https://".length());
	}

	/**
	 * 获取图片流。
	 *
	 * @param path
	 * @return
	 */
	public static InputStream getImageStream(String path) {
		try {
			if (path==null || path.length()==0) {
				throw new IllegalArgumentException("图片路径不能为空");
			}
			if (isHttpUrl(path)) {
				return downloadImageAsStream(path);
			}
			else if (path.startsWith("resources:")) {
				path = path.replace("resources:", "");
				return PathHandler.getInputStreamFromResource(path);
			}
			else {
				// 存储文件的物理路径，兼容 Windows 盘符路径、UNC 路径和相对路径中的中文文件名。
				boolean isAbsolutePath = PathHandler.isAbsolutePath(path);
				String fileAbsolutePath = isAbsolutePath ? path : PathHandler.getProjectPath() + File.separator + path;
				return new FileInputStream(fileAbsolutePath);
			}
		} catch (Exception e) {
			throw new RuntimeException("读取图片失败：" + path, e);
		}
	}

}
