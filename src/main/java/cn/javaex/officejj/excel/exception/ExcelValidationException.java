package cn.javaex.officejj.excel.exception;

/**
 * Excel校验异常
 * 
 * @author 陈霓清
 */
public class ExcelValidationException extends Exception {

	private static final long serialVersionUID = 1L;
	
	private String message;    // 异常信息

	public ExcelValidationException(String message) {
		super();
		this.message = message;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}
	
}
