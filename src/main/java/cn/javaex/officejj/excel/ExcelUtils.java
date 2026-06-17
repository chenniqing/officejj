package cn.javaex.officejj.excel;

import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.Future;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellRangeAddress;

import cn.javaex.officejj.common.util.PathHandler;
import cn.javaex.officejj.common.util.PropertyHandler;
import cn.javaex.officejj.excel.annotation.ExcelSheet;
import cn.javaex.officejj.excel.entity.ExcelImportErrorMarkSetting;
import cn.javaex.officejj.excel.entity.ExcelImportResult;
import cn.javaex.officejj.excel.entity.ExcelSetting;
import cn.javaex.officejj.excel.entity.ExcelSaxReadResult;
import cn.javaex.officejj.excel.entity.ExcelSaxReadSetting;
import cn.javaex.officejj.excel.entity.SheetPrintArea;
import cn.javaex.officejj.excel.exception.ExcelValidationException;
import cn.javaex.officejj.excel.function.ExcelCellHandler;
import cn.javaex.officejj.excel.function.ExcelReadCancelChecker;
import cn.javaex.officejj.excel.function.ExcelReadProgressListener;
import cn.javaex.officejj.excel.function.ExcelImportProcessor;
import cn.javaex.officejj.excel.function.ExcelSaxRowHandler;
import cn.javaex.officejj.excel.help.CellHelper;
import cn.javaex.officejj.excel.help.ExcelImportErrorMarkHelper;
import cn.javaex.officejj.excel.help.ExcelSaxReadHelper;
import cn.javaex.officejj.excel.help.ExcelTextHelper;
import cn.javaex.officejj.excel.help.RowHelper;
import cn.javaex.officejj.excel.help.SheetAnnotationHelper;
import cn.javaex.officejj.excel.help.SheetHelper;
import cn.javaex.officejj.excel.help.SheetMergeHelper;
import cn.javaex.officejj.excel.help.SheetReadHelper;
import cn.javaex.officejj.excel.help.SheetSettingHelper;
import cn.javaex.officejj.excel.help.SheetTemplateHelper;
import cn.javaex.officejj.excel.help.WorkbookHelpler;

/**
 * Excel工具类
 *
 * @author 陈霓清
 */
public class ExcelUtils {

	/** 多线程导出默认分片数。POI 写 Workbook/Sheet 不是线程安全的，真正写入时仍需要串行保护。 */
	private static final int WRITE_EXCEL_THREAD_COUNT = 20;

	// ==================== ↓↓↓↓↓ 获取Excel模板 ↓↓↓↓↓ ====================
	/**
	 * 通过路径读取Excel
	 * @param filePath     例如：D:\\123.xlsx
	 * @return
	 * @throws FileNotFoundException
	 */
	public static Workbook getExcel(String filePath) throws FileNotFoundException {
		return getExcel(new FileInputStream(filePath), true);
	}

	/**
	 * 读取resources文件夹下的Excel
	 * @param filePath     resources文件夹下的路径，例如：template/excel/模板.xlsx
	 * @return
	 * @throws IOException
	 */
	public static Workbook getExcelFromResource(String filePath) throws IOException {
		InputStream in = PathHandler.getInputStreamFromResource(filePath);
		return getExcel(in, true);
	}

	/**
	 * 通过流读取Excel，不关闭调用方传入的输入流。
	 * @param in
	 * @return
	 * @throws Exception
	 */
	public static Workbook getExcel(InputStream in) {
		return getExcel(in, false);
	}

	/**
	 * 通过流读取Excel。
	 * 工具类内部创建的输入流需要关闭；调用方传入的输入流默认由调用方自己关闭。
	 * @param in
	 * @param closeInputStream 是否在读取完成后关闭输入流
	 * @return
	 */
	private static Workbook getExcel(InputStream in, boolean closeInputStream) {
		if (in==null) {
			throw new IllegalArgumentException("Excel输入流不能为空");
		}
		ZipSecureFile.setMinInflateRatio(0.0);

		try {
			return WorkbookFactory.create(in);
		} catch (Exception e) {
			throw new RuntimeException("读取Excel失败", e);
		} finally {
			if (closeInputStream) {
				IOUtils.closeQuietly(in);
			}
		}
	}
	// ==================== ↑↑↑↑↑ 获取Excel模板 ↑↑↑↑↑ ====================


	// ==================== ↓↓↓↓↓ 读取内容 ↓↓↓↓↓ ====================
	/**
	 * 获取单元格内容
	 * @param cell
	 * @return
	 */
	public static String getCellValue(Cell cell) {
		if (cell == null) {
			return "";
		}

		String cellValue = "";

		switch (cell.getCellType()) {
			case STRING :
				cellValue = cell.getRichStringCellValue().getString().trim();
				break;
			case NUMERIC :
				// 判断是否为日期类型
				if (DateUtil.isCellDateFormatted(cell)) {
					// 用于转化为日期格式
					Date date = cell.getDateCellValue();
					cellValue = String.valueOf(date.getTime());
				} else {
					// 格式化数字
					if (cell.toString().endsWith(".0")) {
						DecimalFormat df = new DecimalFormat("#");
						cellValue = df.format(cell.getNumericCellValue()).toString();
					} else if (cell.toString().contains("E")) {
						DecimalFormat df = new DecimalFormat("#");
						cellValue = df.format(cell.getNumericCellValue()).toString();
					} else {
						cellValue = String.valueOf(cell.getNumericCellValue());
					}
				}
				break;
			case BOOLEAN :
				cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
				break;
			case FORMULA :    // 公式
				switch (cell.getCachedFormulaResultType()) {
	                case NUMERIC:
	                	cellValue = String.valueOf(cell.getNumericCellValue());
	                	break;
	                case STRING:
	                	cellValue = cell.getRichStringCellValue().getString();
	                	break;
	                case BOOLEAN:
	                	cellValue = String.valueOf(cell.getBooleanCellValue());
	                	break;
	                case ERROR:
	                	cellValue = "";    // 公式错误
	                	break;
	                default:
	                	cellValue = "";
	                	break;
	            }
				break;
			case BLANK :
				cellValue = "";
				break;
			case ERROR :
				cellValue = "";
				break;
			default :
				cellValue = "";
		}

		return cellValue;
	}

	/**
	 * 读取Excel的第一个Sheet页，并将每一行转为自定义实体对象
	 * @param in
	 * @param clazz    自定义实体类
	 * @param rowNum   从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(InputStream in, Class<T> clazz, int rowNum) throws Exception {
		return readExcel(in, clazz, 1, rowNum);
	}

	/**
	 * 读取Excel，并将每一行转为自定义实体对象
	 * @param in
	 * @param clazz      自定义实体类
	 * @param sheetNum   读取第几个Sheet页（从1开始计算）
	 * @param rowNum     从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(InputStream in, Class<T> clazz, int sheetNum, int rowNum) throws Exception {
		if (in==null) {
			throw new IllegalArgumentException("Excel输入流不能为空");
		}
		ZipSecureFile.setMinInflateRatio(0.0);

		try (Workbook wb = WorkbookFactory.create(in)) {
			Sheet sheet = wb.getSheetAt(sheetNum-1);
			return readExcel(sheet, clazz, rowNum);
		} finally {
			IOUtils.closeQuietly(in);
		}
	}

	/**
	 * 读取Excel，并将每一行转为自定义实体对象
	 * @param in
	 * @param clazz      自定义实体类
	 * @param sheetName  读取哪一个Sheet页，填写Sheet页名称
	 * @param rowNum     从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(InputStream in, Class<T> clazz, String sheetName, int rowNum) throws Exception {
		if (in==null) {
			throw new IllegalArgumentException("Excel输入流不能为空");
		}
		ZipSecureFile.setMinInflateRatio(0.0);

		try (Workbook wb = WorkbookFactory.create(in)) {
			Sheet sheet = wb.getSheet(sheetName);
			if (sheet==null) {
				throw new IllegalArgumentException("Sheet不存在：" + sheetName);
			}
			return readExcel(sheet, clazz, rowNum);
		} finally {
			IOUtils.closeQuietly(in);
		}
	}

	/**
	 * 读取Excel的第一个Sheet页，并将每一行转为自定义实体对象
	 * @param workbook
	 * @param clazz      自定义实体类
	 * @param rowNum     从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(Workbook workbook, Class<T> clazz, int rowNum) throws Exception {
		if (workbook==null) {
			throw new IllegalArgumentException("Workbook不能为空");
		}
		Sheet sheet = workbook.getSheetAt(0);
		return readExcel(sheet, clazz, rowNum);
	}

	/**
	 * 读取Excel，并将每一行转为自定义实体对象
	 * @param <T>
	 * @param sheet
	 * @param clazz
	 * @param rowNum     从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(Sheet sheet, Class<T> clazz, int rowNum) throws Exception {
		if (sheet==null) {
			throw new IllegalArgumentException("Sheet不能为空");
		}
		return new SheetReadHelper().read(sheet, clazz, rowNum-1);
	}

	/**
	 * 读取Excel第一个Sheet页中指定单元格的内容
	 * @param workbook
	 * @param rowNum      第几个行（从1开始计算）
	 * @param colNum      第几个列（从1开始计算）
	 */
	public static String readExcel(Workbook workbook, int rowNum, int colNum) {
		Sheet sheet = workbook.getSheetAt(0);
		return readExcel(sheet, rowNum, colNum);
	}

	/**
	 * 读取Excel指定单元格的内容
	 * @param sheet
	 * @param rowNum      第几个行（从1开始计算）
	 * @param colNum      第几个列（从1开始计算）
	 */
	public static String readExcel(Sheet sheet, int rowNum, int colNum) {
		Row row = sheet.getRow(rowNum-1);
		if (row==null) {
			return "";
		}

		return getCellValue(row.getCell(colNum-1));
	}

	/**
	 * 获取所有Sheet对象
	 * @since 5.0.0
	 *
	 * @param workbook
	 * @return
	 */
	public static List<Sheet> getAllSheets(Workbook workbook) {
		return IntStream.range(0, workbook.getNumberOfSheets())
			    .mapToObj(workbook::getSheetAt)
			    .collect(Collectors.toList());
	}

	/**
	 * 使用SAX模式低内存读取xlsx。
	 * 适合大文件导入，按批次回调数据，不会创建完整Workbook。
	 * @param in 输入流
	 * @param setting 读取配置
	 * @param rowHandler 批次行处理器
	 * @return
	 * @throws Exception
	 */
	public static ExcelSaxReadResult readExcelBySax(InputStream in, ExcelSaxReadSetting setting, ExcelSaxRowHandler rowHandler) throws Exception {
		return readExcelBySax(in, setting, rowHandler, null, null);
	}

	/**
	 * 使用SAX模式低内存读取xlsx。
	 * 适合导入任务接入进度条、取消任务和最大行数保护。
	 * @param in 输入流
	 * @param setting 读取配置
	 * @param rowHandler 批次行处理器
	 * @param progressListener 进度监听器，允许为空
	 * @param cancelChecker 取消检查器，允许为空
	 * @return
	 * @throws Exception
	 */
	public static ExcelSaxReadResult readExcelBySax(InputStream in, ExcelSaxReadSetting setting, ExcelSaxRowHandler rowHandler,
			ExcelReadProgressListener progressListener, ExcelReadCancelChecker cancelChecker) throws Exception {
		return new ExcelSaxReadHelper().read(in, setting, rowHandler, progressListener, cancelChecker);
	}

	/**
	 * 读取CSV。
	 * @param in 输入流
	 * @return
	 * @throws Exception
	 */
	public static List<String[]> readCsv(InputStream in) throws Exception {
		return readCsv(in, StandardCharsets.UTF_8);
	}

	/**
	 * 读取CSV。
	 * @param in 输入流
	 * @param charset 字符集
	 * @return
	 * @throws Exception
	 */
	public static List<String[]> readCsv(InputStream in, Charset charset) throws Exception {
		return new ExcelTextHelper().readCsv(in, charset);
	}

	/**
	 * 写出CSV。
	 * @param rowList 行数据
	 * @param out 输出流
	 * @throws Exception
	 */
	public static void writeCsv(List<String[]> rowList, OutputStream out) throws Exception {
		writeCsv(rowList, out, StandardCharsets.UTF_8);
	}

	/**
	 * 写出CSV。
	 * @param rowList 行数据
	 * @param out 输出流
	 * @param charset 字符集
	 * @throws Exception
	 */
	public static void writeCsv(List<String[]> rowList, OutputStream out, Charset charset) throws Exception {
		new ExcelTextHelper().writeCsv(rowList, out, charset);
	}

	/**
	 * 将Sheet转成简单HTML table。
	 * @param sheet Sheet对象
	 * @return
	 */
	public static String excelToHtml(Sheet sheet) {
		return new ExcelTextHelper().toHtml(sheet);
	}

	/**
	 * 将Workbook中指定Sheet转成简单HTML table。
	 * @param workbook Workbook对象
	 * @param sheetNum 第几个Sheet，从1开始
	 * @return
	 */
	public static String excelToHtml(Workbook workbook, int sheetNum) {
		if (workbook==null) {
			throw new IllegalArgumentException("Workbook不能为空");
		}
		return excelToHtml(workbook.getSheetAt(sheetNum - 1));
	}

	/**
	 * 将简单HTML table转成Workbook。
	 * @param html HTML table文本
	 * @return
	 */
	public static Workbook htmlToExcel(String html) {
		return new ExcelTextHelper().htmlToWorkbook(html, null);
	}

	// ==================== ↑↑↑↑↑ 读取内容 ↑↑↑↑↑ ====================


	// ==================== ↓↓↓↓↓ 导入错误标记 / 导入任务 ↓↓↓↓↓ ====================
	/**
	 * 标记导入失败行。
	 * @param originalFileBytes 原始导入文件字节
	 * @param rowErrorMap       失败行号和错误信息，行号从1开始
	 * @return
	 */
	public static Workbook markImportErrorRows(byte[] originalFileBytes, Map<Integer, List<String>> rowErrorMap) {
		return markImportErrorRows(originalFileBytes, rowErrorMap, null);
	}

	/**
	 * 标记导入失败行。
	 * @param originalFileBytes 原始导入文件字节
	 * @param rowErrorMap       失败行号和错误信息，行号从1开始
	 * @param setting           错误标记配置
	 * @return
	 */
	public static Workbook markImportErrorRows(byte[] originalFileBytes, Map<Integer, List<String>> rowErrorMap, ExcelImportErrorMarkSetting setting) {
		if (originalFileBytes==null || originalFileBytes.length==0) {
			throw new IllegalArgumentException("原始导入文件不能为空");
		}

		try {
			Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(originalFileBytes));
			return markImportErrorRows(workbook, rowErrorMap, setting);
		} catch (Exception e) {
			throw new RuntimeException("生成错误标记Excel失败", e);
		}
	}

	/**
	 * 标记导入失败行。
	 * 该方法不关闭调用方传入的输入流。
	 * @param in          原始导入文件输入流
	 * @param rowErrorMap 失败行号和错误信息，行号从1开始
	 * @return
	 */
	public static Workbook markImportErrorRows(InputStream in, Map<Integer, List<String>> rowErrorMap) {
		return markImportErrorRows(in, rowErrorMap, null);
	}

	/**
	 * 标记导入失败行。
	 * 该方法不关闭调用方传入的输入流。
	 * @param in          原始导入文件输入流
	 * @param rowErrorMap 失败行号和错误信息，行号从1开始
	 * @param setting     错误标记配置
	 * @return
	 */
	public static Workbook markImportErrorRows(InputStream in, Map<Integer, List<String>> rowErrorMap, ExcelImportErrorMarkSetting setting) {
		if (in==null) {
			throw new IllegalArgumentException("Excel输入流不能为空");
		}

		try {
			Workbook workbook = WorkbookFactory.create(in);
			return markImportErrorRows(workbook, rowErrorMap, setting);
		} catch (Exception e) {
			throw new RuntimeException("生成错误标记Excel失败", e);
		}
	}

	/**
	 * 标记导入失败行。
	 * @param workbook    原始导入Workbook
	 * @param rowErrorMap 失败行号和错误信息，行号从1开始
	 * @return
	 */
	public static Workbook markImportErrorRows(Workbook workbook, Map<Integer, List<String>> rowErrorMap) {
		return markImportErrorRows(workbook, rowErrorMap, null);
	}

	/**
	 * 标记导入失败行。
	 * @param workbook    原始导入Workbook
	 * @param rowErrorMap 失败行号和错误信息，行号从1开始
	 * @param setting     错误标记配置
	 * @return
	 */
	public static Workbook markImportErrorRows(Workbook workbook, Map<Integer, List<String>> rowErrorMap, ExcelImportErrorMarkSetting setting) {
		return new ExcelImportErrorMarkHelper().markErrorRows(workbook, rowErrorMap, setting);
	}

	/**
	 * 标记导入失败行，并直接返回可落库或上传文件服务的字节数组。
	 * @param originalFileBytes 原始导入文件字节
	 * @param rowErrorMap       失败行号和错误信息，行号从1开始
	 * @return
	 */
	public static byte[] markImportErrorRowsToBytes(byte[] originalFileBytes, Map<Integer, List<String>> rowErrorMap) {
		return markImportErrorRowsToBytes(originalFileBytes, rowErrorMap, null);
	}

	/**
	 * 标记导入失败行，并直接返回可落库或上传文件服务的字节数组。
	 * @param originalFileBytes 原始导入文件字节
	 * @param rowErrorMap       失败行号和错误信息，行号从1开始
	 * @param setting           错误标记配置
	 * @return
	 */
	public static byte[] markImportErrorRowsToBytes(byte[] originalFileBytes, Map<Integer, List<String>> rowErrorMap, ExcelImportErrorMarkSetting setting) {
		Workbook workbook = null;
		try {
			workbook = markImportErrorRows(originalFileBytes, rowErrorMap, setting);
			return toByteArray(workbook);
		} catch (Exception e) {
			throw new RuntimeException("生成错误标记Excel失败", e);
		} finally {
			if (workbook!=null) {
				try { workbook.close(); } catch (Exception ignore) {}
			}
		}
	}

	/**
	 * 同步执行导入任务。
	 * 方法内部会把异常转换成失败结果，方便业务系统统一更新任务中心状态。
	 * @param originalFileBytes 原始导入文件字节
	 * @param processor         业务导入处理器
	 * @param <T>
	 * @return
	 */
	public static <T> ExcelImportResult<T> importExcel(byte[] originalFileBytes, ExcelImportProcessor<T> processor) {
		return importExcel(originalFileBytes, processor, null);
	}

	/**
	 * 同步执行导入任务。
	 * 方法内部会把异常转换成失败结果，方便业务系统统一更新任务中心状态。
	 * @param originalFileBytes 原始导入文件字节
	 * @param processor         业务导入处理器
	 * @param setting           错误标记配置
	 * @param <T>
	 * @return
	 */
	public static <T> ExcelImportResult<T> importExcel(byte[] originalFileBytes, ExcelImportProcessor<T> processor, ExcelImportErrorMarkSetting setting) {
		Workbook workbook = null;
		try {
			if (originalFileBytes==null || originalFileBytes.length==0) {
				throw new IllegalArgumentException("原始导入文件不能为空");
			}
			if (processor==null) {
				throw new IllegalArgumentException("Excel导入处理器不能为空");
			}

			ZipSecureFile.setMinInflateRatio(0.0);
			workbook = WorkbookFactory.create(new ByteArrayInputStream(originalFileBytes));
			ExcelImportResult<T> result = processor.process(workbook);
			if (result==null) {
				result = ExcelImportResult.success();
			}
			return completeImportResult(originalFileBytes, result, setting);
		} catch (ExcelValidationException e) {
			ExcelImportResult<T> result = ExcelImportResult.failure(e.getMessage());
			result.setException(e);
			result.setRowErrorMap(e.getRowErrorMap());
			return completeImportResult(originalFileBytes, result, setting);
		} catch (Exception e) {
			ExcelImportResult<T> result = ExcelImportResult.failure("Excel导入失败：" + e.getMessage());
			result.setException(e);
			return result;
		} finally {
			if (workbook!=null) {
				try { workbook.close(); } catch (Exception ignore) {}
			}
		}
	}

	/**
	 * 补全导入结果，存在失败行时自动生成错误文件。
	 * @param originalFileBytes 原始导入文件字节
	 * @param result            导入结果
	 * @param setting           错误标记配置
	 * @param <T>
	 * @return
	 */
	private static <T> ExcelImportResult<T> completeImportResult(byte[] originalFileBytes, ExcelImportResult<T> result, ExcelImportErrorMarkSetting setting) {
		if (result.hasError()) {
			result.setSuccess(false);
			if (result.getFailCount()<=0) {
				result.setFailCount(result.getRowErrorMap().size());
			}
			if (result.getMessage()==null || result.getMessage().length()==0) {
				result.setMessage("导入数据存在错误，请下载错误文件核对");
			}
			result.setErrorFileBytes(markImportErrorRowsToBytes(originalFileBytes, result.getRowErrorMap(), setting));
		}
		return result;
	}
	// ==================== ↑↑↑↑↑ 导入错误标记 / 导入任务 ↑↑↑↑↑ ====================


	// ==================== ↓↓↓↓↓ 写入内容  ↓↓↓↓↓ ====================
	/**
	 * 创建默认 Workbook。
	 * @return
	 */
	public static Workbook createWorkbook() {
		return createWorkbook(0);
	}

	/**
	 * 根据导出数据量创建 Workbook，数据量较大时自动使用 SXSSFWorkbook。
	 * @param size 导出数据量
	 * @return
	 */
	public static Workbook createWorkbook(int size) {
		return new WorkbookHelpler().createWorkbook(size);
	}

	/**
	 * 向指定单元格写入内容
	 * @since 5.0.0
	 *
	 * @param cell
	 * @param content
	 * @throws Exception
	 */
	public static void writeCell(Cell cell, Object content) {
		CellHelper cellHelper = new CellHelper();
		cellHelper.setValue(cell, content);
	}

	/**
	 * 向指定单元格写入内容
	 * @param sheet
	 * @param rowNum      第几个行（从1开始计算）
	 * @param colNum      第几个列（从1开始计算）
	 * @param content     写入内容
	 */
	public static void writeExcel(Sheet sheet, int rowNum, int colNum, String content) {
		if (content==null) {
			return;
		}

		Row row = sheet.getRow(rowNum-1);
		if (row==null) {
			row = sheet.createRow(rowNum-1);
		}

		Cell cell = row.getCell(colNum-1);
		if (cell==null) {
			cell = row.createCell(colNum-1);
		}

		cell.setCellValue(content);
	}

	/**
	 * 设置单元格批注。
	 * @param cell 单元格
	 * @param commentText 批注内容
	 * @param author 批注作者，允许为空
	 */
	public static void setComment(Cell cell, String commentText, String author) {
		if (cell==null) {
			throw new IllegalArgumentException("单元格不能为空");
		}
		if (commentText==null) {
			commentText = "";
		}
		Sheet sheet = cell.getSheet();
		Workbook workbook = sheet.getWorkbook();
		CreationHelper creationHelper = workbook.getCreationHelper();
		Drawing<?> drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = creationHelper.createClientAnchor();
		anchor.setCol1(cell.getColumnIndex());
		anchor.setCol2(cell.getColumnIndex() + 3);
		anchor.setRow1(cell.getRowIndex());
		anchor.setRow2(cell.getRowIndex() + 4);
		Comment comment = drawing.createCellComment(anchor);
		comment.setString(creationHelper.createRichTextString(commentText));
		if (author!=null) {
			comment.setAuthor(author);
		}
		cell.setCellComment(comment);
	}

	/**
	 * 设置超链接。
	 * @param cell 单元格
	 * @param address 链接地址
	 * @param label 单元格显示文本，允许为空
	 */
	public static void setHyperlink(Cell cell, String address, String label) {
		if (cell==null) {
			throw new IllegalArgumentException("单元格不能为空");
		}
		if (address==null || address.trim().length()==0) {
			throw new IllegalArgumentException("超链接地址不能为空");
		}
		CreationHelper creationHelper = cell.getSheet().getWorkbook().getCreationHelper();
		Hyperlink hyperlink = creationHelper.createHyperlink(HyperlinkType.URL);
		hyperlink.setAddress(address);
		cell.setHyperlink(hyperlink);
		if (label!=null) {
			cell.setCellValue(label);
		}
	}

	/**
	 * 冻结窗格。
	 * @param sheet Sheet对象
	 * @param colSplit 冻结列数
	 * @param rowSplit 冻结行数
	 */
	public static void freezePane(Sheet sheet, int colSplit, int rowSplit) {
		if (sheet==null) {
			throw new IllegalArgumentException("Sheet不能为空");
		}
		sheet.createFreezePane(Math.max(0, colSplit), Math.max(0, rowSplit));
	}

	/**
	 * 设置自动筛选。
	 * @param sheet Sheet对象
	 * @param firstRow 起始行，从1开始
	 * @param lastRow 结束行，从1开始
	 * @param firstCol 起始列，从1开始
	 * @param lastCol 结束列，从1开始
	 */
	public static void setAutoFilter(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		if (sheet==null) {
			throw new IllegalArgumentException("Sheet不能为空");
		}
		sheet.setAutoFilter(new CellRangeAddress(firstRow - 1, lastRow - 1, firstCol - 1, lastCol - 1));
	}

	/**
	 * 设置文本包含条件格式。
	 * @param sheet Sheet对象
	 * @param rangeAddress 区域，例如 A2:D100
	 * @param text 需要匹配的文本
	 */
	public static void setTextContainsCondition(Sheet sheet, String rangeAddress, String text) {
		if (sheet==null) {
			throw new IllegalArgumentException("Sheet不能为空");
		}
		if (rangeAddress==null || rangeAddress.trim().length()==0) {
			throw new IllegalArgumentException("条件格式区域不能为空");
		}
		SheetConditionalFormatting formatting = sheet.getSheetConditionalFormatting();
		ConditionalFormattingRule rule = formatting.createConditionalFormattingRule("NOT(ISERROR(SEARCH(\"" + text.replace("\"", "\"\"") + "\"," + rangeAddress.split(":")[0] + ")))");
		PatternFormatting patternFormatting = rule.createPatternFormatting();
		patternFormatting.setFillBackgroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
		patternFormatting.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
		formatting.addConditionalFormatting(new CellRangeAddress[] {CellRangeAddress.valueOf(rangeAddress)}, rule);
	}

	/**
	 * 横向写入列表数据。
	 * 适用于按月份、按周、按图片序号等从左到右扩展的模板区域。
	 * @param sheet Sheet对象
	 * @param rowNum 写入行号，从1开始
	 * @param startColNum 起始列号，从1开始
	 * @param list 列表数据
	 * @param property 属性路径；为空时直接写入列表元素本身
	 */
	public static void writeHorizontalList(Sheet sheet, int rowNum, int startColNum, List<?> list, String property) {
		if (sheet==null) {
			throw new IllegalArgumentException("Sheet不能为空");
		}
		if (rowNum<=0 || startColNum<=0) {
			throw new IllegalArgumentException("行号和列号必须从1开始");
		}
		if (list==null || list.isEmpty()) {
			return;
		}

		Row row = sheet.getRow(rowNum - 1);
		if (row==null) {
			row = sheet.createRow(rowNum - 1);
		}
		CellHelper cellHelper = new CellHelper();
		for (int i=0; i<list.size(); i++) {
			Cell cell = row.getCell(startColNum - 1 + i);
			if (cell==null) {
				cell = row.createCell(startColNum - 1 + i);
			}
			Object value = property==null || property.length()==0 ? list.get(i) : PropertyHandler.getValue(list.get(i), property);
			cellHelper.setValue(cell, value);
		}
	}

	/**
	 * 遍历Sheet中已经存在的单元格并执行回调。
	 * 该方法适合在写出完成后统一追加批注、超链接、样式、条件处理等逻辑。
	 * @param sheet Sheet对象
	 * @param handler 单元格处理器
	 * @throws Exception
	 */
	public static void handleCells(Sheet sheet, ExcelCellHandler handler) throws Exception {
		if (sheet==null) {
			throw new IllegalArgumentException("Sheet不能为空");
		}
		if (handler==null) {
			throw new IllegalArgumentException("单元格处理器不能为空");
		}
		for (Row row : sheet) {
			if (row==null) {
				continue;
			}
			for (Cell cell : row) {
				if (cell!=null) {
					handler.handle(sheet, row, cell);
				}
			}
		}
	}

	/**
	 * 替换Excel中的占位符内容
	 * @param workbook
	 * @param sheetNum      第几个Sheet（从1开始计算）
	 * @param param
	 */
	public static void writeExcel(Workbook workbook, int sheetNum, Map<String, Object> param) {
		writeExcel(workbook.getSheetAt(sheetNum - 1), param);
	}

	/**
	 * 替换Excel中的占位符内容。
	 * @param workbook
	 * @param sheetNum       第几个Sheet（从1开始计算）
	 * @param param
	 * @param autoDataHeight 是否根据文本内容自动调整数据行高
	 */
	public static void writeExcel(Workbook workbook, int sheetNum, Map<String, Object> param, boolean autoDataHeight) {
		writeExcel(workbook.getSheetAt(sheetNum - 1), param, autoDataHeight);
	}

	/**
	 * 替换Excel中的占位符内容
	 * @param sheet
	 * @param param
	 */
	public static void writeExcel(Sheet sheet, Map<String, Object> param) {
		SheetHelper sheetHelper = new SheetTemplateHelper();
		sheetHelper.write(sheet, param);
	}

	/**
	 * 替换Excel中的占位符内容。
	 * 默认不自动调整行高，避免打印模板被撑破；需要长文本自适应时显式传true。
	 * @param sheet
	 * @param param
	 * @param autoDataHeight 是否根据文本内容自动调整数据行高
	 */
	public static void writeExcel(Sheet sheet, Map<String, Object> param, boolean autoDataHeight) {
		SheetHelper sheetHelper = new SheetTemplateHelper();
		sheetHelper.write(sheet, param, autoDataHeight);
	}

	/**
	 * 写入list数据，并返回Workbook对象
	 * @param excelSetting
	 * @throws Exception
	 */
	public static Workbook writeExcel(ExcelSetting excelSetting) throws Exception {
		return writeExcel(null, excelSetting);
	}

	/**
	 * 写入list数据，并返回Workbook对象
	 * @param workbook
	 * @param excelSetting
	 * @return
	 */
	public static Workbook writeExcel(Workbook workbook, ExcelSetting excelSetting) {
		SheetHelper sheetHelper = new SheetSettingHelper();

		// 1.创建 Excel
		if (workbook == null) {
			List<String[]> dataList = excelSetting.getDataList();
			int size = dataList==null ? 0 : dataList.size();
			workbook = new WorkbookHelpler().createWorkbook(size);
		}

		// 2.创建sheet
		Sheet sheet = workbook.createSheet(excelSetting.getSheetName());

		// 3.创建内容
		sheetHelper.write(sheet, excelSetting);

		return workbook;
	}

	/**
	 * 根据注解方式写入list数据，并返回Workbook对象
	 * @param clazz 数据库查询得到的vo实体对象
	 * @param list  数据库查询得到的vo实体对象的数据集合
	 * @return
	 * @throws Exception
	 */
	public static Workbook writeExcel(Class<?> clazz, List<?> list) throws Exception {
		// 设置sheet名称
		String sheetName = SheetHelper.SHEET_NAME;
		String sheetTitle = null;
		ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
		if (excelSheet!=null) {
			sheetName = excelSheet.name();
			sheetTitle = excelSheet.title();
		}

		return writeExcel(null, clazz, list, sheetName, sheetTitle);
	}

	/**
	 * 根据注解方式写入list数据，并返回Workbook对象
	 * @param workbook
	 * @param clazz      数据库查询得到的vo实体对象
	 * @param list       数据库查询得到的vo实体对象的数据集合
	 * @param sheetName  追加创建的Sheet页名称
	 * @return
	 * @throws Exception
	 */
	public static Workbook writeExcel(Workbook workbook, Class<?> clazz, List<?> list, String sheetName) throws Exception {
		return writeExcel(workbook, clazz, list, sheetName, null);
	}

	/**
	 * 根据注解方式写入list数据，并返回Workbook对象
	 * @param workbook    Workbook对象，允许为NULL
	 * @param clazz       数据库查询得到的vo实体对象
	 * @param list        数据库查询得到的vo实体对象的数据集合
	 * @param sheetName   追加创建的Sheet页名称
	 * @param title       追加创建的Sheet页顶部标题
	 * @return
	 * @throws Exception
	 */
	public static Workbook writeExcel(Workbook workbook, Class<?> clazz, List<?> list, String sheetName, String title) throws Exception {
		SheetHelper sheetHelper = new SheetAnnotationHelper();

		// 1.创建 Excel
		if (workbook == null) {
			int size = list==null ? 0 : list.size();
			workbook = new WorkbookHelpler().createWorkbook(size);
		}

		// 2.创建sheet
		Sheet sheet = workbook.createSheet(sheetName);

		// 3.创建内容
		sheetHelper.write(sheet, clazz, list, title);

		return workbook;
	}

	/**
	 * 追加写数据
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @return
	 * @throws Exception
	 */
	public static void writeExcel(Sheet sheet, Class<?> clazz, List<?> list) throws Exception {
		SheetHelper sheetHelper = new SheetAnnotationHelper();
		sheetHelper.write(sheet, clazz, list, null);
	}

	/**
	 * 根据注解方式写入list数据，并返回Workbook对象
	 * 多线程
	 * @param clazz 数据库查询得到的vo实体对象
	 * @param list  数据库查询得到的vo实体对象的数据集合
	 * @return
	 * @throws Exception
	 */
	public static Workbook writeExcelByThreads(Class<?> clazz, List<?> list) throws Exception {
		// 设置sheet名称
		String sheetName = SheetHelper.SHEET_NAME;
		String sheetTitle = null;
		ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
		if (excelSheet!=null) {
			sheetName = excelSheet.name();
			sheetTitle = excelSheet.title();
		}

		return writeExcelByThreads(null, clazz, list, sheetName, sheetTitle);
	}

	/**
	 * 根据注解方式写入list数据，并返回Workbook对象
	 * 多线程
	 * @param workbook
	 * @param clazz      数据库查询得到的vo实体对象
	 * @param list       数据库查询得到的vo实体对象的数据集合
	 * @param sheetName  追加创建的Sheet页名称
	 * @return
	 * @throws Exception
	 */
	public static Workbook writeExcelByThreads(Workbook workbook, Class<?> clazz, List<?> list, String sheetName) throws Exception {
		return writeExcelByThreads(workbook, clazz, list, sheetName, null);
	}

	/**
	 * 根据注解方式写入list数据，并返回Workbook对象
	 * 多线程
	 * @param workbook    Workbook对象，允许为NULL
	 * @param clazz       数据库查询得到的vo实体对象
	 * @param list        数据库查询得到的vo实体对象的数据集合
	 * @param sheetName   追加创建的Sheet页名称
	 * @param title       追加创建的Sheet页顶部标题
	 * @return
	 * @throws Exception
	 */
	public static Workbook writeExcelByThreads(Workbook workbook, Class<?> clazz, List<?> list, String sheetName, String title) throws Exception {
		SheetAnnotationHelper sheetHelper = new SheetAnnotationHelper();

		// 1.创建 Excel
		if (workbook == null) {
			int size = list==null ? 0 : list.size();
			workbook = new WorkbookHelpler().createWorkbook(size);
		}

		// 2.创建sheet
		Sheet sheet = workbook.createSheet(sheetName);

		// 3.多线程写入数据；空数据也要保留表头，行为与普通 writeExcel 保持一致
		int size = list==null ? 0 : list.size();
		if (size <= WRITE_EXCEL_THREAD_COUNT) {
			sheetHelper.write(sheet, clazz, list, title);
		} else {
			sheetHelper.setHeader(sheet, clazz, title);
			int startRowIndex = sheet.getLastRowNum() + 1;
			writeDataByThreads(sheetHelper, sheet, clazz, list, startRowIndex);
		}

		return workbook;
	}

	/**
	 * 追加写数据
	 * 多线程
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @return
	 * @throws Exception
	 */
	public static void writeExcelByThreads(Sheet sheet, Class<?> clazz, List<?> list) throws Exception {
		SheetAnnotationHelper sheetHelper = new SheetAnnotationHelper();

		int size = list==null ? 0 : list.size();
		if (size <= WRITE_EXCEL_THREAD_COUNT) {
			sheetHelper.write(sheet, clazz, list, null);
		} else {
			// 空 Sheet 先创建表头，非空 Sheet 只补基础列宽/格式缓存，然后从末尾继续追加
			if (sheet.getRow(0)==null) {
				sheetHelper.setHeader(sheet, clazz, null);
			} else {
				sheetHelper.setBasicData(sheet, clazz);
			}
			int startRowIndex = sheet.getLastRowNum() + 1;
			writeDataByThreads(sheetHelper, sheet, clazz, list, startRowIndex);
		}
	}

	/**
	 * 按固定起始行分片写入数据。
	 * 注意：POI 的 Workbook/Sheet 写操作不是线程安全的，所以真正写入时需要对 Sheet 加锁；
	 * 每个分片使用预先计算好的起始行，避免旧实现按抢锁顺序追加导致数据乱序。
	 * @param sheetHelper
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param startRowIndex
	 * @throws Exception
	 */
	private static void writeDataByThreads(SheetAnnotationHelper sheetHelper, Sheet sheet, Class<?> clazz, List<?> list, int startRowIndex) throws Exception {
		if (list==null || list.isEmpty()) {
			return;
		}

		int size = list.size();
		int nThreads = Math.min(WRITE_EXCEL_THREAD_COUNT, size);
		int chunkSize = (size + nThreads - 1) / nThreads;
		ExecutorService executorService = Executors.newFixedThreadPool(nThreads);
		List<Future<Void>> futures = new ArrayList<Future<Void>>(nThreads);

		for (int fromIndex=0; fromIndex<size; fromIndex+=chunkSize) {
			int toIndex = Math.min(fromIndex + chunkSize, size);
			final int rowIndex = startRowIndex + fromIndex;
			final List<?> listThreads = list.subList(fromIndex, toIndex);

			Callable<Void> task = () -> {
				synchronized (sheet) {
					sheetHelper.createData(sheet, clazz, listThreads, rowIndex, false);
				}
				return null;
			};
			futures.add(executorService.submit(task));
		}
		executorService.shutdown();

		for (Future<Void> future : futures) {
			try {
				future.get();
			} catch (InterruptedException e) {
				executorService.shutdownNow();
				Thread.currentThread().interrupt();
				throw new RuntimeException("多线程导出被中断", e);
			} catch (ExecutionException e) {
				executorService.shutdownNow();
				Throwable cause = e.getCause();
				if (cause instanceof Exception) {
					throw (Exception) cause;
				}
				throw new RuntimeException("多线程导出失败", cause);
			}
		}

		// 所有分片写完后再统一合并，避免跨分片的同组数据被拆开。
		synchronized (sheet) {
			sheetHelper.mergeDataRows(sheet, clazz, startRowIndex, startRowIndex + size - 1);
			sheetHelper.applyAutoColumnWidth(sheet, clazz);
		}
	}
	// ==================== ↑↑↑↑↑ 写入内容 ↑↑↑↑↑ ====================


	// ==================== ↓↓↓↓↓ 设置单元格内容 ↓↓↓↓↓ ====================
	/**
	 * 自动合并列
	 * @since 5.0.0
	 *
	 * @param sheet
	 * @param colNum      第几列（从1开始计算）
	 * @param firstRow    起始行（从1开始计算）
	 * @param lastRow     终止行（从1开始计算，允许为null）
	 */
	public static void setAutoMergeCol(Sheet sheet, int colNum, int firstRow, Integer lastRow) {
		if (lastRow == null) {
			lastRow = sheet.getLastRowNum() + 1;
		}

		SheetHelper sheetHelper = new SheetMergeHelper();
		sheetHelper.setAutoMergeCol(sheet, colNum-1, firstRow-1, lastRow-1, null);
	}

	/**
	 * 自动合并列，并用指定列作为合并边界。
	 * 例如第2列班主任按第1列班级拆分合并：setAutoMergeCol(sheet, 2, 2, null, 1)。
	 * @since 5.1.3
	 *
	 * @param sheet
	 * @param colNum             第几列（从1开始计算）
	 * @param firstRow           起始行（从1开始计算）
	 * @param lastRow            终止行（从1开始计算，允许为null）
	 * @param mergeByColNumArr   依赖列（从1开始计算），当前列值和依赖列值都相同才合并
	 */
	public static void setAutoMergeCol(Sheet sheet, int colNum, int firstRow, Integer lastRow, int... mergeByColNumArr) {
		setAutoMergeCol(sheet, colNum, firstRow, lastRow, null, mergeByColNumArr);
	}

	/**
	 * 自动合并列（自定义样式）
	 * @since 5.0.0
	 *
	 * @param sheet
	 * @param colNum      第几列（从1开始计算）
	 * @param firstRow    起始行（从1开始计算）
	 * @param lastRow     终止行（从1开始计算，允许为null）
	 * @param clazz       创建自定义样式类，例如传参：MyCellStyle.class
	 */
	public static void setAutoMergeCol(Sheet sheet, int colNum, int firstRow, Integer lastRow, Class<?> clazz) {
		if (lastRow == null) {
			lastRow = sheet.getLastRowNum() + 1;
		}

		SheetHelper sheetHelper = new SheetMergeHelper();
		sheetHelper.setAutoMergeCol(sheet, colNum-1, firstRow-1, lastRow-1, clazz);
	}

	/**
	 * 自动合并列（自定义样式），并用指定列作为合并边界。
	 * @since 5.1.3
	 *
	 * @param sheet
	 * @param colNum             第几列（从1开始计算）
	 * @param firstRow           起始行（从1开始计算）
	 * @param lastRow            终止行（从1开始计算，允许为null）
	 * @param clazz              创建自定义样式类，例如传参：MyCellStyle.class
	 * @param mergeByColNumArr   依赖列（从1开始计算），当前列值和依赖列值都相同才合并
	 */
	public static void setAutoMergeCol(Sheet sheet, int colNum, int firstRow, Integer lastRow, Class<?> clazz, int... mergeByColNumArr) {
		if (lastRow == null) {
			lastRow = sheet.getLastRowNum() + 1;
		}

		SheetHelper sheetHelper = new SheetMergeHelper();
		sheetHelper.setAutoMergeCol(sheet, colNum-1, firstRow-1, lastRow-1, clazz, toZeroBasedColumns(mergeByColNumArr));
	}

	/**
	 * 自动合并行
	 * @since 5.0.0
	 *
	 * @param sheet
	 * @param rowNum      第几行（从1开始计算）
	 * @param firstCol    起始列（从1开始计算）
	 * @param lastCol     终止列（从1开始计算，允许为null）
	 */
	public static void setAutoMergeRow(Sheet sheet, int rowNum, int firstCol, Integer lastCol) {
		if (lastCol == null) {
			Row row = sheet.getRow(rowNum - 1);
			if (row==null) {
				throw new IllegalArgumentException("自动合并行失败，行不存在：" + rowNum);
			}
			lastCol = (int) row.getLastCellNum();
		}

		SheetHelper sheetHelper = new SheetMergeHelper();
		sheetHelper.setAutoMergeRow(sheet, rowNum-1, firstCol-1, lastCol-1, null);
	}

	/**
	 * 自动合并行（自定义样式）
	 * @since 5.0.0
	 *
	 * @param sheet
	 * @param rowNum      第几行（从1开始计算）
	 * @param firstCol    起始列（从1开始计算）
	 * @param lastCol     终止列（从1开始计算，允许为null）
	 * @param clazz       创建自定义样式类，例如传参：MyCellStyle.class
	 */
	public static void setAutoMergeRow(Sheet sheet, int rowNum, int firstCol, Integer lastCol, Class<?> clazz) {
		if (lastCol == null) {
			Row row = sheet.getRow(rowNum - 1);
			if (row==null) {
				throw new IllegalArgumentException("自动合并行失败，行不存在：" + rowNum);
			}
			lastCol = (int) row.getLastCellNum();
		}

		SheetHelper sheetHelper = new SheetMergeHelper();
		sheetHelper.setAutoMergeRow(sheet, rowNum-1, firstCol-1, lastCol-1, clazz);
	}

	/**
	 * 设置合并单元格
	 * @param sheet
	 * @param firstRow    起始行（从1开始计算）
	 * @param lastRow     终止行（从1开始计算）
	 * @param firstCol    起始列（从1开始计算）
	 * @param lastCol     终止列（从1开始计算）
	 */
	public static void setMerge(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		SheetHelper sheetHelper = new SheetMergeHelper();
		sheetHelper.setMerge(sheet, firstRow-1, lastRow-1, firstCol-1, lastCol-1, null);
	}

	/**
	 * 设置合并单元格（自定义样式）
	 * @param sheet
	 * @param firstRow    起始行（从1开始计算）
	 * @param lastRow     终止行（从1开始计算）
	 * @param firstCol    起始列（从1开始计算）
	 * @param lastCol     终止列（从1开始计算）
	 * @param clazz       创建自定义样式类，例如传参：MyCellStyle.class
	 */
	public static void setMerge(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, Class<?> clazz) {
		SheetHelper sheetHelper = new SheetMergeHelper();
		sheetHelper.setMerge(sheet, firstRow-1, lastRow-1, firstCol-1, lastCol-1, clazz);
	}

	/**
	 * 画线条
	 * @since 5.0.0
	 *
	 * @param row1 起点行，从1开始
	 * @param col1 起点列，从1开始
	 * @param row2 终点行，从1开始
	 * @param col2 终点列，从1开始
	 * @param colorRGB 颜色， int数组{r, g, b} 例如：new int[]{0, 0, 0}
	 * @param shrinkRatio 收缩率，例如：填0.8，以防止遮挡单元格内容
	 * @param lineWidth 粗细，单位为磅 例如：2.0
	 * @param isDottedLine 是否是虚线，false为实线
	 */
	public static void drawLine(Sheet sheet, int row1, int col1, int row2, int col2,
			int[] colorRGB, double lineWidth, double shrinkRatio, boolean isDottedLine) {
		if (shrinkRatio >= 1 || shrinkRatio <= 0) {
			new SheetHelper().drawLine(sheet, col1-1, row1-1, col2-1, row2-1, colorRGB, lineWidth, false, isDottedLine);
		} else {
			new SheetHelper().drawLine(sheet, col1-1, row1-1, col2-1, row2-1, colorRGB, lineWidth, shrinkRatio, false, isDottedLine);
		}
	}

	/**
	 * 画线条（带箭头）
	 * @since 5.0.0
	 *
	 * @param row1 起点行，从1开始
	 * @param col1 起点列，从1开始
	 * @param row2 终点行，从1开始
	 * @param col2 终点列，从1开始
	 * @param colorRGB 颜色， int数组{r, g, b} 例如：new int[]{0, 0, 0}
	 * @param shrinkRatio 收缩率，例如：填0.8，以防止遮挡单元格内容
	 * @param lineWidth 粗细，单位为磅 例如：2.0
	 * @param isDottedLine 是否是虚线，false为实线
	 */
	public static void drawLineWithArrow(Sheet sheet, int row1, int col1, int row2, int col2,
			int[] colorRGB, double lineWidth, double shrinkRatio, boolean isDottedLine) {
		if (shrinkRatio >= 1 || shrinkRatio <= 0) {
			new SheetHelper().drawLine(sheet, col1-1, row1-1, col2-1, row2-1, colorRGB, lineWidth, true, isDottedLine);
		} else {
			new SheetHelper().drawLine(sheet, col1-1, row1-1, col2-1, row2-1, colorRGB, lineWidth, shrinkRatio, true, isDottedLine);
		}
	}

	/**
	 * 刷新公式计算
	 * 设置了单元格的值，但公式单元格不会自动更新，使用此方法手动触发公式评估
	 * @since 5.0.0
	 *
	 * @param workbook
	 */
	public static void refreshFormula(Workbook workbook) {
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

		for (Sheet sheet : workbook) {
			for (Row row : sheet) {
				for (Cell cell : row) {
					if (cell.getCellType() == CellType.FORMULA) {
						evaluator.evaluateFormulaCell(cell);
					}
				}
			}
		}
	}

	/**
	 * 刷新公式计算
	 * 设置了单元格的值，但公式单元格不会自动更新，使用此方法手动触发公式评估
	 * @since 5.0.0
	 *
	 * @param workbook
	 * @param sheet
	 */
	public static void refreshFormula(Workbook workbook, Sheet sheet) {
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.FORMULA) {
					evaluator.evaluateFormulaCell(cell);
				}
			}
		}
	}

	/**
	 * 设置下拉选项（第一个Sheet页）
	 * @param workbook
	 * @param colNum          第几个列（从1开始计算）
	 * @param startRow        第几个行设置开始（从1开始计算）
	 * @param endRow          第几个行设置结束（从1开始计算）
	 * @param selectDataArr   下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public static void setSelect(Workbook workbook, int colNum, int startRow, int endRow, String[] selectDataArr) {
		Sheet sheet = workbook.getSheetAt(0);
		setSelect(sheet, colNum, startRow, endRow, selectDataArr);
	}

	/**
	 * 设置下拉选项（指定Sheet页）
	 * @param workbook
	 * @param sheetName       Sheet页名称
	 * @param colNum          第几个列（从1开始计算）
	 * @param startRow        第几个行设置开始（从1开始计算）
	 * @param endRow          第几个行设置结束（从1开始计算）
	 * @param selectDataArr   下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public static void setSelect(Workbook workbook, String sheetName, int colNum, int startRow, int endRow, String[] selectDataArr) {
		Sheet sheet = workbook.getSheet(sheetName);
		if (sheet==null) {
			throw new IllegalArgumentException("Sheet不存在：" + sheetName);
		}
		setSelect(sheet, colNum, startRow, endRow, selectDataArr);
	}

	/**
	 * 设置下拉选项
	 * @param sheet
	 * @param colNum          第几个列（从1开始计算）
	 * @param startRow        第几个行设置开始（从1开始计算）
	 * @param endRow          第几个行设置结束（从1开始计算）
	 * @param selectDataArr   下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public static void setSelect(Sheet sheet, int colNum, int startRow, int endRow, String[] selectDataArr) {
		new SheetHelper().setSelect(sheet, colNum-1, startRow-1, endRow-1, selectDataArr);
	}

	/**
	 * 设置只读（随机密码）
	 * @param workbook
	 */
	public static void setReadOnly(Workbook workbook) {
		new WorkbookHelpler().setReadOnly(workbook, null);
	}

	/**
	 * 设置只读
	 * @param workbook
	 * @param password    密码
	 */
	public static void setReadOnly(Workbook workbook, String password) {
		new WorkbookHelpler().setReadOnly(workbook, password);
	}

	/**
	 * 设置只读（随机密码）
	 * @param sheet
	 * @param colNumArr   哪几列设置只读（从1开始计算）  int[] colNumArr = new int[]{1, 2};
	 */
	public static void setReadOnly(Sheet sheet, int[] colNumArr) {
		new SheetHelper().setReadOnly(sheet, null);
		new SheetHelper().setEditable(sheet, colNumArr);
	}

	/**
	 * 设置只读
	 * @param sheet
	 * @param colNumArr   哪几列设置只读（从1开始计算）  int[] colNumArr = new int[]{1, 2};
	 * @param password    密码
	 */
	public static void setReadOnly(Sheet sheet, int[] colNumArr, String password) {
		new SheetHelper().setReadOnly(sheet, password);
		new SheetHelper().setEditable(sheet, colNumArr);
	}

	/**
	 * 设置固定列、固定行（第一个Sheet页）
	 * @param workbook
	 * @param colNum      前N列固定（从1开始计算）
	 * @param rowNum      前N行固定（从1开始计算）
	 */
	public static void setFixed(Workbook workbook, int colNum, int rowNum) {
		Sheet sheet = workbook.getSheetAt(0);
		setFixed(sheet, colNum, rowNum);
	}

	/**
	 * 设置固定列、固定行
	 * @param sheet
	 * @param colNum      前N列固定（从1开始计算）
	 * @param rowNum      前N行固定（从1开始计算）
	 */
	public static void setFixed(Sheet sheet, int colNum, int rowNum) {
		sheet.createFreezePane(colNum, rowNum, colNum, rowNum);
	}

	/**
	 * 设置单元格样式
	 * @param cell
	 * @param clazz
	 */
	public static void setCellStyle(Cell cell, Class<?> clazz) {
		new CellHelper().setCellStyle(cell, clazz);
	}
	// ==================== ↑↑↑↑↑ 设置单元格内容 ↑↑↑↑↑ ====================


	// ==================== ↓↓↓↓↓ 设置模板内容 ↓↓↓↓↓ ====================
	/**
	 * 复制行
	 * @since 5.0.0
	 *
	 * @param sheet
	 * @param srcStartRow 从1开始计算
	 */
	public static void copyRow(Sheet sheet, int srcStartRow) {
		Row sourceRow = sheet.getRow(srcStartRow - 1);
		Row targetRow = sheet.createRow(srcStartRow);

		new RowHelper().copyRow(sheet.getWorkbook(), sourceRow, targetRow);
	}

	/**
	 * 插入行
	 * @since 5.0.0
	 *
	 * @param sheet
	 * @param startRow 起始行（从1开始计算）
	 * @param rows     插入多少行
	 */
	public static void insertRow(Sheet sheet, int startRow, int rows) {
		new SheetHelper().insertRow(sheet, startRow - 1, rows);
	}

	/**
	 * 复制模板
	 * @since 5.0.0
	 *
	 * @param sheet
	 * @param templateRows          模板行数
	 * @param copyTimes             复制几次
	 * @param makePageBreakByBlock  是否设置分页（打印区域）
	 */
	public static void copyTemplate(Sheet sheet, int templateRows, int copyTimes, boolean makePageBreakByBlock) {
		SheetHelper sheetHelper = new SheetTemplateHelper();
		sheetHelper.copyTemplate(sheet, templateRows, copyTimes, makePageBreakByBlock);
	}

	/**
	 * 复制Sheet
	 * @since 5.0.0
	 *
	 * @param workbook
	 * @param sourceSheetName    源Sheet名称（例如：JES1）
	 * @param targetSheetName    目标Sheet名称（例如：JES）
	 * @param copyCount          复制次数（例如：如果复制2次，则会生成JES2、JES3）
	 */
	public static void copySheets(Workbook workbook, String sourceSheetName, String targetSheetName, int copyCount) {
		SheetHelper sheetHelper = new SheetTemplateHelper();
		sheetHelper.copySheets(workbook, sourceSheetName, targetSheetName, copyCount, null);
	}

	/**
	 * 复制Sheet（设置打印区域）
	 * @since 5.0.0
	 *
	 * @param workbook
	 * @param sourceSheetName    源Sheet名称（例如：JES1）
	 * @param targetSheetName    目标Sheet名称（例如：JES）
	 * @param copyCount          复制次数（例如：如果复制2次，则会生成JES2、JES3）
	 * @param printFirstRow      打印区域起始行，（从1开始计算）
	 * @param printLastRow       打印区域终止行，（从1开始计算）
	 * @param printFirstColumn   打印区域起始列，如"A"
	 * @param printLastColumn    打印区域终止列，如"F"
	 */
	public static void copySheets(Workbook workbook, String sourceSheetName, String targetSheetName, int copyCount,
			int printFirstRow, int printLastRow, String printFirstColumn, String printLastColumn) {
		SheetPrintArea printArea = new SheetPrintArea(printFirstRow, printLastRow, printFirstColumn, printLastColumn);

		SheetHelper sheetHelper = new SheetTemplateHelper();
		sheetHelper.copySheets(workbook, sourceSheetName, targetSheetName, copyCount, printArea);
	}
	// ==================== ↑↑↑↑↑ 设置模板内容 ↑↑↑↑↑ ====================


	// ==================== ↓↓↓↓↓ 输出Excel文件 ↓↓↓↓↓ ====================
	/**
	 * 输出Excel到指定路径，默认写完后关闭Workbook。
	 * @param workbook
	 * @param filePath       文件写到哪里的全路径，例如：D:\\1.xlsx
	 */
	public static void output(Workbook workbook, String filePath) {
		output(workbook, filePath, true);
	}

	/**
	 * 输出Excel到指定路径。
	 * 默认方法会关闭Workbook；需要继续复用Workbook时，可调用本重载并传false。
	 * @param workbook
	 * @param filePath       文件写到哪里的全路径，例如：D:\\1.xlsx
	 * @param closeWorkbook  输出后是否关闭Workbook
	 */
	public static void output(Workbook workbook, String filePath, boolean closeWorkbook) {
		// 保证这个文件的父文件夹必须要存在
		File targetFile = new File(filePath);
		File parentFile = targetFile.getParentFile();
		if (parentFile!=null && !parentFile.exists()) {
			parentFile.mkdirs();
		}

		FileOutputStream out = null;
		try {
			out = new FileOutputStream(targetFile);
			workbook.write(out);
			out.flush();
		} catch (Exception e) {
			throw new RuntimeException("输出Excel失败：" + filePath, e);
		} finally {
			if (out != null) {
				try { out.close(); } catch (IOException e) {}
			}
			if (closeWorkbook && workbook!=null) {
				try { workbook.close(); } catch (Exception ignore) {}
			}
		}
	}

	/**
	 * 输出Excel字节到指定路径。
	 * 适合保存导入任务生成的错误文件。
	 * @param fileBytes 文件字节
	 * @param filePath  文件写到哪里的全路径，例如：D:\\1.xlsx
	 */
	public static void output(byte[] fileBytes, String filePath) {
		if (fileBytes==null || fileBytes.length==0) {
			throw new IllegalArgumentException("Excel文件字节不能为空");
		}

		File targetFile = new File(filePath);
		File parentFile = targetFile.getParentFile();
		if (parentFile!=null && !parentFile.exists()) {
			parentFile.mkdirs();
		}

		FileOutputStream out = null;
		try {
			out = new FileOutputStream(targetFile);
			out.write(fileBytes);
			out.flush();
		} catch (Exception e) {
			throw new RuntimeException("输出Excel失败：" + filePath, e);
		} finally {
			if (out != null) {
				try { out.close(); } catch (IOException e) {}
			}
		}
	}

	/**
	 * 下载Excel（兼容 javax 和 jakarta Servlet 环境），默认写完后关闭Workbook。
	 * @param response    javax.servlet.http.HttpServletResponse 或 jakarta.servlet.http.HttpServletResponse
	 * @param workbook
	 * @param filename    带后缀的文件名，例如："test.xlsx"
	 * @throws IOException
	 */
	public static void download(Object response, Workbook workbook, String filename) throws IOException {
		download(response, workbook, filename, true);
	}

	/**
	 * 下载Excel（兼容 javax 和 jakarta Servlet 环境）。
	 * 默认方法会关闭Workbook；需要继续复用Workbook时，可调用本重载并传false。
	 * @param response    javax.servlet.http.HttpServletResponse 或 jakarta.servlet.http.HttpServletResponse
	 * @param workbook
	 * @param filename    带后缀的文件名，例如："test.xlsx"
	 * @param closeWorkbook 下载后是否关闭Workbook
	 * @throws IOException
	 */
	public static void download(Object response, Workbook workbook, String filename, boolean closeWorkbook) throws IOException {
		OutputStream out = null;
		try {
			// 用反射调用setContentType和setHeader
			Method setContentType = response.getClass().getMethod("setContentType", String.class);
			Method setHeader = response.getClass().getMethod("setHeader", String.class, String.class);
			Method getOutputStream = response.getClass().getMethod("getOutputStream");

			setContentType.setAccessible(true);
			setHeader.setAccessible(true);
			getOutputStream.setAccessible(true);

			setContentType.invoke(response, "application/octet-stream; charset=utf-8");
			String encodedFilename = java.net.URLEncoder.encode(filename, "UTF-8").replaceAll("\\+", "%20");
			setHeader.invoke(response, "Content-Disposition", "attachment; filename*=UTF-8''" + encodedFilename);

			out = new BufferedOutputStream((OutputStream) getOutputStream.invoke(response));
			workbook.write(out);
			out.flush();
		} catch (Exception e) {
			throw new IOException("Download excel failed", e);
		} finally {
			if (out != null) {
				try { out.close(); } catch (IOException e) {}
			}
			if (closeWorkbook && workbook!=null) {
				try { workbook.close(); } catch (Exception ignore) {}
			}
		}
	}

	/**
	 * 下载Excel字节（兼容 javax 和 jakarta Servlet 环境）。
	 * 适合下载导入任务生成的错误文件。
	 * @param response    javax.servlet.http.HttpServletResponse 或 jakarta.servlet.http.HttpServletResponse
	 * @param fileBytes   文件字节
	 * @param filename    带后缀的文件名，例如："导入错误.xlsx"
	 * @throws IOException
	 */
	public static void download(Object response, byte[] fileBytes, String filename) throws IOException {
		if (fileBytes==null || fileBytes.length==0) {
			throw new IllegalArgumentException("Excel文件字节不能为空");
		}

		OutputStream out = null;
		try {
			// 用反射调用setContentType和setHeader，兼容 javax.servlet 和 jakarta.servlet。
			Method setContentType = response.getClass().getMethod("setContentType", String.class);
			Method setHeader = response.getClass().getMethod("setHeader", String.class, String.class);
			Method getOutputStream = response.getClass().getMethod("getOutputStream");

			setContentType.setAccessible(true);
			setHeader.setAccessible(true);
			getOutputStream.setAccessible(true);

			setContentType.invoke(response, "application/octet-stream; charset=utf-8");
			String encodedFilename = java.net.URLEncoder.encode(filename, "UTF-8").replaceAll("\\+", "%20");
			setHeader.invoke(response, "Content-Disposition", "attachment; filename*=UTF-8''" + encodedFilename);

			out = new BufferedOutputStream((OutputStream) getOutputStream.invoke(response));
			out.write(fileBytes);
			out.flush();
		} catch (Exception e) {
			throw new IOException("Download excel failed", e);
		} finally {
			if (out != null) {
				try { out.close(); } catch (IOException e) {}
			}
		}
	}

	/**
	 * 将Workbook写入字节数组，不关闭传入的Workbook，便于调用方自行决定生命周期。
	 * @param workbook
	 * @return
	 * @throws IOException
	 */
	public static byte[] toByteArray(Workbook workbook) throws IOException {
		try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
			workbook.write(bos);
			return bos.toByteArray();
		}
	}
	// ==================== ↑↑↑↑↑ 输出Excel文件 ↑↑↑↑↑ ====================


	// ==================== ↓↓↓↓↓ 其他辅助方法 ↓↓↓↓↓ ====================
	/**
	 * 数字转列名
	 * @param colNum    第几列（从1开始计算，即1=A）
	 * @return               列名（如A、AB等）
	 */
	public static String getColName(int colNum) {
		return CellReference.convertNumToColString(colNum - 1);
	}

	/**
	 * 列名转数字
	 * @param colName    列名（如A、AB等）
	 * @return           列索引（从1开始，即A=1）
	 */
	public static int getColNum(String colName) {
		return CellReference.convertColStringToIndex(colName) + 1;
	}

	/**
	 * 把面向用户的1-based列号转换为POI内部使用的0-based列号。
	 * @param colNumArr
	 * @return
	 */
	private static int[] toZeroBasedColumns(int... colNumArr) {
		if (colNumArr==null || colNumArr.length==0) {
			return new int[0];
		}

		int[] result = new int[colNumArr.length];
		for (int i=0; i<colNumArr.length; i++) {
			if (colNumArr[i]<1) {
				throw new IllegalArgumentException("列号必须从1开始：" + colNumArr[i]);
			}
			result[i] = colNumArr[i] - 1;
		}

		return result;
	}
	// ==================== ↑↑↑↑↑ 其他辅助方法 ↑↑↑↑↑ ====================

}
