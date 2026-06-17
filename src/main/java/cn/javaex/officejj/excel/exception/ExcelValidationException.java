package cn.javaex.officejj.excel.exception;

import java.util.List;
import java.util.Map;

/**
 * Excel校验异常
 *
 * @author 陈霓清
 */
public class ExcelValidationException extends Exception {

	private static final long serialVersionUID = 1L;

	private String message;    // 异常信息
	private Map<Integer, List<String>> rowErrorMap;    // 失败行号和错误信息，行号从1开始

	public ExcelValidationException(String message) {
		super(message);
		this.message = message;
	}

	public ExcelValidationException(String message, Map<Integer, List<String>> rowErrorMap) {
		super(message);
		this.message = message;
		this.rowErrorMap = rowErrorMap;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}

	public Map<Integer, List<String>> getRowErrorMap() {
		return rowErrorMap;
	}

	public void setRowErrorMap(Map<Integer, List<String>> rowErrorMap) {
		this.rowErrorMap = rowErrorMap;
	}

}
