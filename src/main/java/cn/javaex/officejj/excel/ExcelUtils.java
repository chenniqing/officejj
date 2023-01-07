package cn.javaex.officejj.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
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

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import cn.javaex.officejj.common.util.FileHandler;
import cn.javaex.officejj.common.util.PathHandler;
import cn.javaex.officejj.excel.annotation.ExcelSheet;
import cn.javaex.officejj.excel.entity.ExcelSetting;
import cn.javaex.officejj.excel.help.PreviewHelper;
import cn.javaex.officejj.excel.help.SheetAnnotationHelper;
import cn.javaex.officejj.excel.help.SheetHelper;
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
	
	/**
	 * 获取单元格内容
	 * @param cell
	 * @return
	 */
	public static String getCellValue(Cell cell) {
		if (cell==null) {
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
				try {
					cellValue = String.valueOf(cell.getNumericCellValue());
				} catch (IllegalStateException e) {
					cellValue = String.valueOf(cell.getRichStringCellValue());
				} catch (Exception e) {
					cellValue = cell.getCellFormula();
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
	
	/**
	 * 导出Excel
	 * @param excelSetting
	 * @throws Exception
	 */
	public static Workbook writeExcel(ExcelSetting excelSetting) throws Exception {
		return writeExcel(null, excelSetting);
	}
	
	/**
	 * 导出Excel
	 * @param wb
	 * @param excelSetting
	 * @return
	 */
	public static Workbook writeExcel(Workbook wb, ExcelSetting excelSetting) {
		SheetHelper sheetHelper = new SheetSettingHelper();
		
		// 1.0 创建 Excel
		if (wb==null) {
			List<String[]> dataList = excelSetting.getDataList();
			int size = dataList==null ? 0 : dataList.size();
			wb = new WorkbookHelpler().createWorkbook(size);
		}
		
		// 2.0 创建sheet
		Sheet sheet = wb.createSheet(excelSetting.getSheetName());
		
		// 3.0 创建内容
		sheetHelper.write(sheet, excelSetting);
		
		return wb;
	}
	
	/**
	 * 根据注解方式导出Excel
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
	 * 根据注解方式导出Excel
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
	 * 根据注解方式导出Excel（手动指定Sheet页名称）
	 * @param wb         Workbook对象
	 * @param clazz      数据库查询得到的vo实体对象
	 * @param list       数据库查询得到的vo实体对象的数据集合
	 * @param sheetName  追加创建的Sheet页名称
	 * @return
	 * @throws Exception
	 */
	public static Workbook writeExcel(Workbook wb, Class<?> clazz, List<?> list, String sheetName) throws Exception {
		return writeExcel(wb, clazz, list, sheetName, null);
	}
	
	/**
	 * 根据注解方式导出Excel（手动指定Sheet页名称）
	 * 多线程
	 * @param wb         Workbook对象
	 * @param clazz      数据库查询得到的vo实体对象
	 * @param list       数据库查询得到的vo实体对象的数据集合
	 * @param sheetName  追加创建的Sheet页名称
	 * @return
	 * @throws Exception
	 */
	public static Workbook writeExcelByThreads(Workbook wb, Class<?> clazz, List<?> list, String sheetName) throws Exception {
		return writeExcelByThreads(wb, clazz, list, sheetName, null);
	}
	
	/**
	 * 根据注解方式导出Excel（手动指定Sheet页名称）
	 * @param wb          Workbook对象，允许为NULL
	 * @param clazz       数据库查询得到的vo实体对象
	 * @param list        数据库查询得到的vo实体对象的数据集合
	 * @param sheetName   追加创建的Sheet页名称
	 * @param title       追加创建的Sheet页顶部标题
	 * @return
	 * @throws Exception
	 */
	public static Workbook writeExcel(Workbook wb, Class<?> clazz, List<?> list, String sheetName, String title) throws Exception {
		SheetHelper sheetHelper = new SheetAnnotationHelper();
		
		// 1.0 创建 Excel
		if (wb==null) {
			int size = list==null ? 0 : list.size();
			wb = new WorkbookHelpler().createWorkbook(size);
		}
		
		// 2.0 创建sheet
		Sheet sheet = wb.createSheet(sheetName);
		
		// 3.0 创建内容
		sheetHelper.write(sheet, clazz, list, title);
		
		return wb;
	}
	
	/**
	 * 根据注解方式导出Excel（手动指定Sheet页名称）
	 * 多线程
	 * @param wb          Workbook对象，允许为NULL
	 * @param clazz       数据库查询得到的vo实体对象
	 * @param list        数据库查询得到的vo实体对象的数据集合
	 * @param sheetName   追加创建的Sheet页名称
	 * @param title       追加创建的Sheet页顶部标题
	 * @return
	 * @throws Exception
	 */
	public static Workbook writeExcelByThreads(Workbook wb, Class<?> clazz, List<?> list, String sheetName, String title) throws Exception {
		SheetHelper sheetHelper = new SheetAnnotationHelper();
		
		// 1.0 创建 Excel
		if (wb==null) {
			int size = list==null ? 0 : list.size();
			wb = new WorkbookHelpler().createWorkbook(size);
		}
		
		// 2.0 创建sheet
		Sheet sheet = wb.createSheet(sheetName);
		
		// 3.0 多线程写入数据
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
				
			}
			if (yushu != 0) {
				List<?> otherList = list.subList(index2, size);
				sheetHelper.writeByThreads(sheet, clazz, otherList);
			}
		}
		
		return wb;
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
				
			}
			if (yushu != 0) {
				List<?> otherList = list.subList(index2, size);
				sheetHelper.writeByThreads(sheet, clazz, otherList);
			}
		}
	}
	
	/**
	 * 写入Excel单元格
	 * @param wb
	 * @param sheetNum    第几个Sheet页（从1开始计算）
	 * @param rowNum      第几个行（从1开始计算）
	 * @param colNum      第几个列（从1开始计算）
	 * @param content     写入内容
	 */
	public static void writeExcel(Workbook wb, int sheetNum, int rowNum, int colNum, String content) {
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		writeExcel(sheet, rowNum, colNum, content);
	}
	
	/**
	 * 写入Excel单元格
	 * @param wb
	 * @param sheetName   Sheet页名称
	 * @param rowNum      第几个行（从1开始计算）
	 * @param colNum      第几个列（从1开始计算）
	 * @param content     写入内容
	 */
	public static void writeExcel(Workbook wb, String sheetName, int rowNum, int colNum, String content) {
		Sheet sheet = wb.getSheet(sheetName);
		writeExcel(sheet, rowNum, colNum, content);
	}
	
	/**
	 * 写入Excel单元格
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
	 * @param wb
	 * @param clazz      自定义实体类
	 * @param rowNum     从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(Workbook wb, Class<T> clazz, int rowNum) throws Exception {
		Sheet sheet = wb.getSheetAt(0);
		return readExcel(sheet, clazz, rowNum);
	}
	
	/**
	 * 读取Excel，并将每一行转为自定义实体对象
	 * @param wb
	 * @param clazz      自定义实体类
	 * @param sheetName  读取第几个Sheet页（从1开始计算）
	 * @param rowNum     从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(Workbook wb, Class<T> clazz, int sheetNum, int rowNum) throws Exception {
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		return readExcel(sheet, clazz, rowNum);
	}
	
	/**
	 * 读取Excel，并将每一行转为自定义实体对象
	 * @param wb
	 * @param clazz      自定义实体类
	 * @param sheetName  读取哪一个Sheet页，填写Sheet页名称
	 * @param rowNum     从第几行开始读取（从1开始计算）
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readExcel(Workbook wb, Class<T> clazz, String sheetName, int rowNum) throws Exception {
		Sheet sheet = wb.getSheet(sheetName);
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
	 * @param wb
	 * @param rowNum      第几个行（从1开始计算）
	 * @param colNum      第几个列（从1开始计算）
	 */
	public static String readExcel(Workbook wb, int rowNum, int colNum) {
		Sheet sheet = wb.getSheetAt(0);
		return readExcel(sheet, rowNum, colNum);
	}
	
	/**
	 * 读取Excel单元格
	 * @param wb
	 * @param sheetNum    第几个Sheet页（从1开始计算）
	 * @param rowNum      第几个行（从1开始计算）
	 * @param colNum      第几个列（从1开始计算）
	 */
	public static String readExcel(Workbook wb, int sheetNum, int rowNum, int colNum) {
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		return readExcel(sheet, rowNum, colNum);
	}
	
	/**
	 * 读取Excel单元格
	 * @param wb
	 * @param sheetName   Sheet页名称
	 * @param rowNum      第几个行（从1开始计算）
	 * @param colNum      第几个列（从1开始计算）
	 */
	public static String readExcel(Workbook wb, String sheetName, int rowNum, int colNum) {
		Sheet sheet = wb.getSheet(sheetName);
		return readExcel(sheet, rowNum, colNum);
	}

	/**
	 * 读取Excel单元格
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
	 * 设置下拉选项（第一个Sheet页）
	 * @param wb
	 * @param colNum          第几个列（从1开始计算）
	 * @param startRow        第几个行设置开始（从1开始计算）
	 * @param endRow          第几个行设置结束（从1开始计算）
	 * @param selectDataArr   下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public static void setSelect(Workbook wb, int colNum, int startRow, int endRow, String[] selectDataArr) {
		setSelect(wb, 1, colNum, startRow, endRow, selectDataArr);
	}
	
	/**
	 * 设置下拉选项
	 * @param wb
	 * @param sheetNum        第几个Sheet页（从1开始计算）
	 * @param colNum          第几个列（从1开始计算）
	 * @param startRow        第几个行设置开始（从1开始计算）
	 * @param endRow          第几个行设置结束（从1开始计算）
	 * @param selectDataArr   下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public static void setSelect(Workbook wb, int sheetNum, int colNum, int startRow, int endRow, String[] selectDataArr) {
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		setSelect(sheet, colNum, startRow, endRow, selectDataArr);
	}
	
	/**
	 * 设置下拉选项
	 * @param wb
	 * @param sheetName       Sheet页名称
	 * @param colNum          第几个列（从1开始计算）
	 * @param startRow        第几个行设置开始（从1开始计算）
	 * @param endRow          第几个行设置结束（从1开始计算）
	 * @param selectDataArr   下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public static void setSelect(Workbook wb, String sheetName, int colNum, int startRow, int endRow, String[] selectDataArr) {
		Sheet sheet = wb.getSheet(sheetName);
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
	 * 设置合并单元格（第一个Sheet页）
	 * @param wb
	 * @param firstRow    起始行（从1开始计算）
	 * @param lastRow     终止行（从1开始计算）
	 * @param firstCol    起始列（从1开始计算）
	 * @param lastCol     终止列（从1开始计算）
	 */
	public static void setMerge(Workbook wb, int firstRow, int lastRow, int firstCol, int lastCol) {
		setMerge(wb, 1, firstRow, lastRow, firstCol, lastCol);
	}
	
	/**
	 * 设置合并单元格
	 * @param wb
	 * @param sheetNum    第几个Sheet页（从1开始计算）
	 * @param firstRow    起始行（从1开始计算）
	 * @param lastRow     终止行（从1开始计算）
	 * @param firstCol    起始列（从1开始计算）
	 * @param lastCol     终止列（从1开始计算）
	 */
	public static void setMerge(Workbook wb, int sheetNum, int firstRow, int lastRow, int firstCol, int lastCol) {
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		setMerge(sheet, firstRow, lastRow, firstCol, lastCol);
	}
	
	/**
	 * 设置合并单元格
	 * @param wb
	 * @param sheetName   Sheet页名称
	 * @param firstRow    起始行（从1开始计算）
	 * @param lastRow     终止行（从1开始计算）
	 * @param firstCol    起始列（从1开始计算）
	 * @param lastCol     终止列（从1开始计算）
	 */
	public static void setMerge(Workbook wb, String sheetName, int firstRow, int lastRow, int firstCol, int lastCol) {
		Sheet sheet = wb.getSheet(sheetName);
		setMerge(sheet, firstRow, lastRow, firstCol, lastCol);
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
		new SheetHelper().setMerge(sheet, firstRow-1, lastRow-1, firstCol-1, lastCol-1);
	}

	/**
	 * 输出Excel到指定路径
	 * @param wb
	 * @param filePath       文件写到哪里的全路径，例如：D:\\1.xlsx
	 */
	public static void output(Workbook wb, String filePath) {
		// 保证这个文件的父文件夹必须要存在
		File targetFile = new File(filePath);
		if (!targetFile.getParentFile().exists()) {
			targetFile.getParentFile().mkdirs();
		}
		
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(targetFile);
			wb.write(out);
			out.flush();
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(out);
		}
	}
	
	/**
	 * 下载Excel
	 * @param wb
	 * @param filename       文件名，例如：test.xlsx
	 * @throws IOException
	 */
	public static void download(Workbook wb, String filename) throws IOException {
		HttpServletResponse response = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getResponse();
		download(response, wb, filename);
	}
	
	/**
	 * 下载Excel
	 * @param wb
	 * @param filename       文件名，例如：test.xlsx
	 * @throws IOException
	 */
	public static void download(HttpServletResponse response, Workbook wb, String filename) throws IOException {
		String folderPath = PathHandler.getFolderPath();
		
		String fileUrl = folderPath + File.separator + filename;
		
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(fileUrl);
			wb.write(out);
			out.flush();
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(out);
		}
		
		FileHandler.downloadFile(response, fileUrl);
	}
	
	/**
	 * 设置只读（随机密码）
	 * @param word 
	 */
	public static void setReadOnly(Workbook wb) {
		new WorkbookHelpler().setReadOnly(wb, null);
	}
	
	/**
	 * 设置只读
	 * @param wb
	 * @param password    密码
	 */
	public static void setReadOnly(Workbook wb, String password) {
		new WorkbookHelpler().setReadOnly(wb, password);
	}
	
	/**
	 * 设置固定列、固定行
	 * @param wb
	 * @param sheetNum    第几个Sheet页（从1开始计算）
	 * @param colNum      前N列固定（从1开始计算）
	 * @param rowNum      前N行固定（从1开始计算）
	 */
	public static void setFixed(Workbook wb, int sheetNum, int colNum, int rowNum) {
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		setFixed(sheet, colNum, rowNum);
	}

	/**
	 * 设置固定列、固定行
	 * @param wb
	 * @param sheetName   Sheet页名称
	 * @param colNum      前N列固定（从1开始计算）
	 * @param rowNum      前N行固定（从1开始计算）
	 */
	public static void setFixed(Workbook wb, String sheetName, int colNum, int rowNum) {
		Sheet sheet = wb.getSheet(sheetName);
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
	 * 替换Excel中的占位符内容
	 * @param wb
	 * @param sheetNum    第几个Sheet页（从1开始计算）
	 * @param param
	 */
	public static void writeExcel(Workbook wb, int sheetNum, Map<String, Object> param) {
		Sheet sheet = wb.getSheetAt(sheetNum-1);
		writeExcel(sheet, param);
	}

	/**
	 * 替换Excel中的占位符内容
	 * @param wb
	 * @param sheetName   Sheet页名称
	 * @param param
	 */
	public static void writeExcel(Workbook wb, String sheetName, Map<String, Object> param) {
		Sheet sheet = wb.getSheet(sheetName);
		writeExcel(sheet, param);
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
	 * Excel转Html
	 * @param filePath     excel文件路径，例如：D:\\Temp\\1.xlsx
	 * @return             返回生成的html文件的全路径，例如：D:\\Temp\\1_html\\1.html
	 * @throws Exception 
	 */
	public static String excelToHtml(String filePath) throws Exception {
		return new PreviewHelper().excelToHtml(filePath);
	}
	
}
