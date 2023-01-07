package cn.javaex.officejj.word.help;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 帮助类的父类
 * 
 * @author 陈霓清
 */
public class Helper {

	/**
	 * 替换换行符
	 * @param str
	 * @return
	 */
	public String replaceBr(String str) {
		return str.replace("<br>", "<br/>").replace("\r\n", "<br/>").replace("\n", "<br/>");
	}
	
	/**
	 * 提取${xx}中的文本（只提取第一个）
	 * @param str
	 * @return
	 */
	public String getPlaceholder(String str) {
		String placeholder = "";
		
		String patern = "(?<=\\$\\{)[^\\}]+";
		Pattern pattern = Pattern.compile(patern);
		Matcher matcher = pattern.matcher(str);
		while (matcher.find()) {
			placeholder = matcher.group();
			break;
		}
		
		return placeholder;
	}
	
}
