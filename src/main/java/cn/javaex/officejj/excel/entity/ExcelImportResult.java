package cn.javaex.officejj.excel.entity;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel导入结果。
 *
 * @param <T> 导入后需要返回给业务方的数据类型
 * @author 陈霓清
 */
public class ExcelImportResult<T> {
	private boolean success = true;                                                // 是否导入成功
	private String message;                                                        // 任务结果说明
	private int totalCount;                                                        // 总行数
	private int successCount;                                                      // 成功行数
	private int failCount;                                                         // 失败行数
	private List<T> data = new ArrayList<T>();                                     // 业务方需要返回的数据
	private Map<Integer, List<String>> rowErrorMap = new LinkedHashMap<Integer, List<String>>(); // 失败行号和错误信息，行号从1开始
	private byte[] errorFileBytes;                                                 // 标红后的错误文件字节
	private String errorFileName;                                                  // 错误文件名
	private Throwable exception;                                                   // 任务异常，便于业务方记录日志

	/**
	 * 创建成功结果。
	 * @param <T>
	 * @return
	 */
	public static <T> ExcelImportResult<T> success() {
		return new ExcelImportResult<T>();
	}

	/**
	 * 创建失败结果。
	 * @param message
	 * @param <T>
	 * @return
	 */
	public static <T> ExcelImportResult<T> failure(String message) {
		ExcelImportResult<T> result = new ExcelImportResult<T>();
		result.setSuccess(false);
		result.setMessage(message);
		return result;
	}

	/**
	 * 是否存在导入错误。
	 * @return
	 */
	public boolean hasError() {
		return rowErrorMap!=null && !rowErrorMap.isEmpty();
	}

	/**
	 * 是否已经生成错误文件。
	 * @return
	 */
	public boolean hasErrorFile() {
		return errorFileBytes!=null && errorFileBytes.length>0;
	}

	/**
	 * 追加某一行的错误信息。
	 * @param excelRowNum Excel行号，从1开始
	 * @param message 错误信息
	 */
	public void addError(int excelRowNum, String message) {
		if (rowErrorMap==null) {
			rowErrorMap = new LinkedHashMap<Integer, List<String>>();
		}
		List<String> list = rowErrorMap.get(excelRowNum);
		if (list==null) {
			list = new ArrayList<String>();
			rowErrorMap.put(excelRowNum, list);
		}
		list.add(message);
	}

	/**
	 * 追加某一行的多条错误信息。
	 * @param excelRowNum Excel行号，从1开始
	 * @param messageList 错误信息
	 */
	public void addErrors(int excelRowNum, List<String> messageList) {
		if (messageList==null || messageList.isEmpty()) {
			return;
		}
		if (rowErrorMap==null) {
			rowErrorMap = new LinkedHashMap<Integer, List<String>>();
		}
		List<String> list = rowErrorMap.get(excelRowNum);
		if (list==null) {
			list = new ArrayList<String>();
			rowErrorMap.put(excelRowNum, list);
		}
		list.addAll(messageList);
	}

	public boolean isSuccess() {
		return success;
	}

	public void setSuccess(boolean success) {
		this.success = success;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}

	public int getTotalCount() {
		return totalCount;
	}

	public void setTotalCount(int totalCount) {
		this.totalCount = totalCount;
	}

	public int getSuccessCount() {
		return successCount;
	}

	public void setSuccessCount(int successCount) {
		this.successCount = successCount;
	}

	public int getFailCount() {
		return failCount;
	}

	public void setFailCount(int failCount) {
		this.failCount = failCount;
	}

	public List<T> getData() {
		return data;
	}

	public void setData(List<T> data) {
		this.data = data;
	}

	public Map<Integer, List<String>> getRowErrorMap() {
		return rowErrorMap;
	}

	public void setRowErrorMap(Map<Integer, List<String>> rowErrorMap) {
		this.rowErrorMap = rowErrorMap;
	}

	public byte[] getErrorFileBytes() {
		return errorFileBytes;
	}

	public void setErrorFileBytes(byte[] errorFileBytes) {
		this.errorFileBytes = errorFileBytes;
	}

	public String getErrorFileName() {
		return errorFileName;
	}

	public void setErrorFileName(String errorFileName) {
		this.errorFileName = errorFileName;
	}

	public Throwable getException() {
		return exception;
	}

	public void setException(Throwable exception) {
		this.exception = exception;
	}
}
