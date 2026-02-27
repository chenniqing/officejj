package cn.javaex.officejj.excel;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import cn.javaex.officejj.common.util.PathHandler;
import cn.javaex.officejj.excel.annotation.ExcelSheet;
import cn.javaex.officejj.excel.entity.ExcelSetting;
import cn.javaex.officejj.excel.entity.SheetPrintArea;
import cn.javaex.officejj.excel.help.CellHelper;
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

	// ==================== ↓↓↓↓↓ 获取Excel模板 ↓↓↓↓↓ ====================
	/**
	 * 通过路径读取Excel
	 * @param filePath     例如：D:\\123.xlsx
	 * @return
	 * @throws FileNotFoundException 
	 */
	public static Workbook getExcel(String filePath) throws FileNotFoundException {
		return getExcel(new FileInputStream(filePath));
	}
	
	/**
	 * 读取resources文件夹下的Excel
	 * @param filePath     resources文件夹下的路径，例如：template/excel/模板.xlsx
	 * @return
	 * @throws IOException 
	 */
	public static Workbook getExcelFromResource(String filePath) throws IOException {
		InputStream in = PathHandler.getInputStreamFromResource(filePath);
		return getExcel(in);
	}
	
	/**
	 * 通过流读取Excel
	 * @param in
	 * @return
	 * @throws Exception
	 */
	public static Workbook getExcel(InputStream in) {
		Workbook wb = null;
		
		try {
			wb = WorkbookFactory.create(in);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(in);
		}
		
		return wb;
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
		Workbook wb = WorkbookFactory.create(in);
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		
		return readExcel(sheet, clazz, rowNum);
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
		Workbook wb = WorkbookFactory.create(in);
		Sheet sheet = wb.getSheet(sheetName);
		
		return readExcel(sheet, clazz, rowNum);
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
	
	// ==================== ↑↑↑↑↑ 读取内容 ↑↑↑↑↑ ====================
	

	// ==================== ↓↓↓↓↓ 写入内容  ↓↓↓↓↓ ====================
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
		if (content==null || content.length()==0) {
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
	 * 替换Excel中的占位符内容
	 * @param sheet
	 * @param param
	 */
	public static void writeExcel(Sheet sheet, Map<String, Object> param) {
		SheetHelper sheetHelper = new SheetTemplateHelper();
		sheetHelper.write(sheet, param);
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
		SheetHelper sheetHelper = new SheetAnnotationHelper();
		
		// 1.创建 Excel
		if (workbook == null) {
			int size = list==null ? 0 : list.size();
			workbook = new WorkbookHelpler().createWorkbook(size);
		}
		
		// 2.创建sheet
		Sheet sheet = workbook.createSheet(sheetName);
		
		// 3.多线程写入数据
		int nThreads = 20;
		int size = list.size();
		int yushu = size % nThreads;
		if (size <= nThreads) {
			sheetHelper.write(sheet, clazz, list, title);
		} else {
			sheetHelper.setHeader(sheet, clazz, title);
			
			ExecutorService executorService = Executors.newFixedThreadPool(nThreads);
			List<Future<Integer>> futures = new ArrayList<Future<Integer>>(nThreads);
			
			int index1 = 0;
			int index2 = 0;
			for (int i = 0; i < nThreads; i++) {
				index1 = size / nThreads * i;
				index2 = size / nThreads * (i + 1);
				final List<?> listThreads = list.subList(index1, index2);
				
				Callable<Integer> task = () -> {
					sheetHelper.writeByThreads(sheet, clazz, listThreads);
					return 1;
				};
				futures.add(executorService.submit(task));
			}
			executorService.shutdown();
			
			try {
				executorService.awaitTermination(Long.MAX_VALUE, TimeUnit.NANOSECONDS);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
			if (yushu != 0) {
				List<?> otherList = list.subList(index2, size);
				sheetHelper.writeByThreads(sheet, clazz, otherList);
			}
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
		SheetHelper sheetHelper = new SheetAnnotationHelper();
		sheetHelper.setBasicData(sheet, clazz);
		
		int nThreads = 20;
		int size = list.size();
		int yushu = size % nThreads;
		if (size <= nThreads) {
			sheetHelper.write(sheet, clazz, list, null);
		} else {
			ExecutorService executorService = Executors.newFixedThreadPool(nThreads);
			List<Future<Integer>> futures = new ArrayList<Future<Integer>>(nThreads);
			
			int index1 = 0;
			int index2 = 0;
			for (int i = 0; i < nThreads; i++) {
				index1 = size / nThreads * i;
				index2 = size / nThreads * (i + 1);
				final List<?> listThreads = list.subList(index1, index2);
				
				Callable<Integer> task = () -> {
					sheetHelper.writeByThreads(sheet, clazz, listThreads);
					return 1;
				};
				futures.add(executorService.submit(task));
			}
			executorService.shutdown();
			
			try {
				executorService.awaitTermination(Long.MAX_VALUE, TimeUnit.NANOSECONDS);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
			if (yushu != 0) {
				List<?> otherList = list.subList(index2, size);
				sheetHelper.writeByThreads(sheet, clazz, otherList);
			}
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
			lastCol = (int) (sheet.getRow(rowNum - 1)).getLastCellNum();
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
			lastCol = (int) (sheet.getRow(rowNum - 1)).getLastCellNum();
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
	 * 输出Excel到指定路径
	 * @param workbook
	 * @param filePath       文件写到哪里的全路径，例如：D:\\1.xlsx
	 */
	public static void output(Workbook workbook, String filePath) {
		// 保证这个文件的父文件夹必须要存在
		File targetFile = new File(filePath);
		if (!targetFile.getParentFile().exists()) {
			targetFile.getParentFile().mkdirs();
		}
		
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(targetFile);
			workbook.write(out);
			out.flush();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (out != null) {
				try { out.close(); } catch (IOException e) {}
			}
			try { workbook.close(); } catch (Exception ignore) {}
		}
	}
	
	/**
	 * 下载Excel（兼容 javax 和 jakarta Servlet 环境）
	 * @param response    javax.servlet.http.HttpServletResponse 或 jakarta.servlet.http.HttpServletResponse
	 * @param workbook
	 * @param filename    带后缀的文件名，例如："test.xlsx"
	 * @throws IOException
	 */
	public static void download(Object response, Workbook workbook, String filename) throws IOException {
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
			e.printStackTrace();
			throw new IOException("Download excel failed", e);
		} finally {
			if (out != null) {
				try { out.close(); } catch (IOException e) {}
			}
			try { workbook.close(); } catch (Exception ignore) {}
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
	// ==================== ↑↑↑↑↑ 其他辅助方法 ↑↑↑↑↑ ====================
	
}
