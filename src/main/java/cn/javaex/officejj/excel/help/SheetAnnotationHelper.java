package cn.javaex.officejj.excel.help;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import cn.javaex.officejj.common.entity.Picture;
import cn.javaex.officejj.excel.annotation.ExcelCell;
import cn.javaex.officejj.excel.annotation.ExcelStyle;
import cn.javaex.officejj.excel.function.DefaultExcelValueConverter;
import cn.javaex.officejj.excel.function.ExcelValueConverter;
import cn.javaex.officejj.excel.style.DefaultCellStyle;
import cn.javaex.officejj.excel.style.ICellStyle;

/**
 * 注解法导出Excel
 *
 * @author 陈霓清
 */
public class SheetAnnotationHelper extends SheetHelper {
	/** 自动行高最小行数 */
	private static final int MIN_AUTO_HEIGHT_LINES = 1;
	/** 单元格左右内边距折算的字符数，避免文本贴边时估算偏小 */
	private static final double CELL_PADDING_CHAR_WIDTH = 1.0D;
	/** Excel 允许的最大列宽，单位为字符宽度 */
	private static final int MAX_EXCEL_COLUMN_WIDTH = 255;
	/** 自动列宽额外预留的字符宽度，避免内容紧贴单元格边缘 */
	private static final double AUTO_COLUMN_WIDTH_PADDING_CHAR_WIDTH = 2.0D;

	/**
	 * 创建Header
	 * @param sheet
	 * @param clazz
	 * @param title
	 * @throws Exception
	 */
	@Override
	public synchronized void setHeader(Sheet sheet, Class<?> clazz, String title) throws Exception {
		// 当前写到了第几行（从1开始计算）
		int rowNum = sheet.getLastRowNum() < 0 ? 0 : sheet.getLastRowNum();
		Row row = sheet.getRow(0);

		// 1.0 设置基础属性
		this.setBasicData(sheet, clazz);

		// 表示是新建的Sheet页
		if (row==null) {
			// 2.0 设置标题
			if (title!=null && title.length()>0) {
				rowNum = this.createTtile(sheet, clazz, title);
			}

			// 3.0 设置表头
			int headerRows = this.getHeaderRows(clazz);
			if (headerRows==1) {
				// 单行表头
				rowNum = this.createHeader(sheet, clazz, rowNum);
			} else {
				// 多行表头
				rowNum = this.createHeaders(sheet, clazz, rowNum, headerRows);
			}
		}

		// 注解启用自动列宽时，即使只创建表头，也要按表头内容先调整一次列宽。
		this.applyAutoColumnWidth(sheet, clazz);
	}

	/**
	 * 创建内容
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param title
	 * @throws Exception
	 */
	@Override
	public synchronized void write(Sheet sheet, Class<?> clazz, List<?> list, String title) throws Exception {
		// 当前写到了第几行（从1开始计算）
		int rowNum = sheet.getLastRowNum() < 0 ? 0 : sheet.getLastRowNum();
		Row row = sheet.getRow(0);

		// 1.0 设置基础属性
		this.setBasicData(sheet, clazz);

		// 表示是新建的Sheet页
		if (row==null) {
			// 2.0 设置标题
			if (title!=null && title.length()>0) {
				rowNum = this.createTtile(sheet, clazz, title);
			}

			// 3.0 设置表头
			int headerRows = this.getHeaderRows(clazz);
			if (headerRows==1) {
				// 单行表头
				rowNum = this.createHeader(sheet, clazz, rowNum);
			} else {
				// 多行表头
				rowNum = this.createHeaders(sheet, clazz, rowNum, headerRows);
			}
		} else {
			rowNum = rowNum + 1;
		}

		// 4.0 设置数据
		this.createData(sheet, clazz, list, rowNum);

		// 数据写入完成后再统一调整列宽，确保表头和数据内容都会参与计算。
		this.applyAutoColumnWidth(sheet, clazz);
	}

	/**
	 * 创建内容（多线程）
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param title
	 * @throws Exception
	 */
	@Override
	public synchronized void writeByThreads(Sheet sheet, Class<?> clazz, List<?> list) throws Exception {
		// 当前写到了第几行（从1开始计算）
		int rowNum = sheet.getLastRowNum();

		// 设置数据
		this.createData(sheet, clazz, list, rowNum + 1);

		// 直接调用本方法时，也在当前批次写完后刷新自动列宽。
		this.applyAutoColumnWidth(sheet, clazz);
	}

	/**
	 * 设置基础属性
	 * @param sheet
	 * @param clazz
	 */
	@Override
	public void setBasicData(Sheet sheet, Class<?> clazz) {
		int colIndex = 0;    // 列索引
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);

		Field[] fieldArr = clazz.getDeclaredFields();
		for (int i=0; i<fieldArr.length; i++) {
			Field field = fieldArr[i];
			// 设置类的私有属性可访问
			field.setAccessible(true);
			// 得到每一个成员变量上的 ExcelCell 注解
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}
			// 跳过被归入组的列
			if (skipMap.get(field.getName())!=null) {
				continue;
			}

			int sort = excelCell.sort()==0 ? colIndex : (excelCell.sort() - 1);

			// 设置列宽
			if (this.isAutoColumnWidth(excelStyle)) {
				// 自动列宽最终会在数据写完后重新计算；这里先给一个最小起点，避免初始列宽过大。
				sheet.setColumnWidth(sort, this.toColumnWidthUnits(this.resolveMinColumnWidth(excelStyle)));
			} else {
				sheet.setColumnWidth(sort, excelCell.width() * BASE_COLUMN_WIDTH);
			}

			// 设置值替换属性
			String[] replaceArr = excelCell.replace();
			if (replaceArr.length>0) {
				Map<String, String> map = new HashMap<String, String>();
				// {"1_男", "0_女"}
				for (String replace : replaceArr) {
					// 1_男
					String[] arr = replace.split("_", 2);
					if (arr.length==2) {
						map.put(arr[0], arr[1]);
					}
				}

				replaceMap.put(String.valueOf(sort), map);
			}
			// 设置格式化属性
			String format = excelCell.format();
			if (format.length()>0) {
				if (field.getType()==LocalDateTime.class || field.getType()==LocalDate.class) {
					DateTimeFormatter dtf = DateTimeFormatter.ofPattern(format);
					formatMap.put(String.valueOf(sort), dtf);
				}
				else if (field.getType()==Date.class) {
					SimpleDateFormat sdf = new SimpleDateFormat(format);
					formatMap.put(String.valueOf(sort), sdf);
				}
			}

			// 合并组
			int mergeCol = excelCell.group();
			if (mergeCol>1 && excelCell.sort()==0) {
				int num = 0;
				for (int j=(i+1); j<fieldArr.length; j++) {
					Field temp = fieldArr[j];
					temp.setAccessible(true);
					if (temp.getAnnotation(ExcelCell.class)==null) {
						continue;
					}

					skipMap.put(temp.getName(), temp.getName());

					num++;

					if (num==(mergeCol-1)) {
						break;
					}
				}
			}

			colIndex++;
		}
	}

	/**
	 * 根据注解得到要设置的表头行数
	 * @param clazz
	 * @return
	 */
	private int getHeaderRows(Class<?> clazz) {
		int headerRows = 1;

		Field[] fieldArr = clazz.getDeclaredFields();
		for (Field field : fieldArr) {
			// 设置类的私有属性可访问
			field.setAccessible(true);
			// 得到每一个成员变量上的 ExcelCell 注解
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}

			headerRows = Math.max(1, excelCell.name().length);

			break;
		}

		return headerRows;
	}

	/**
	 * 设置标题
	 * @param sheet
	 * @param clazz
	 * @param title
	 * @return              返回当前写到第几行
	 * @throws Exception
	 */
	private int createTtile(Sheet sheet, Class<?> clazz, String title) throws Exception {
		Row row = sheet.createRow(0);

		// 样式
		CellStyle cellStyle = null;
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);
		if (excelStyle==null) {
			cellStyle = new DefaultCellStyle().createTitleStyle(sheet.getWorkbook());
		} else {
			ICellStyle obj = (ICellStyle) Class.forName(excelStyle.cellStyle()).getDeclaredConstructor().newInstance();
			cellStyle = obj.createTitleStyle(sheet.getWorkbook());

			// 行高
			int height = excelStyle.titleHeight();
			if (height>0) {
				row.setHeight((short) (height * BASE_ROW_HEIGHT));
			}
		}

		// 设置单元格
		Cell cell = row.createCell(0);
		cell.setCellValue(title);
		cell.setCellStyle(cellStyle);

		// 得到该类的所有成员变量，计算得到需要合并的单元格
		int length = 0;
		Field[] declaredFields = clazz.getDeclaredFields();
		for (Field field : declaredFields) {
			// 设置类的私有属性可访问
			field.setAccessible(true);

			// 得到每一个成员变量上的 ExcelCell 注解
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}

			length++;

			// 有组合并的话，长度要减少
			int mergeCol = excelCell.group();
			if (mergeCol>1) {
				length = length - (mergeCol-1);
			}
		}

		// 设置合并
		// 四个参数分别是：起始行、终止行、起始列、终止列（从0开始计算）
		if (length>1) {
			CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, length-1);
			sheet.addMergedRegion(cellRangeAddress);
		}

		return 1;
	}

	/**
	 * 设置单行表头
	 * @param sheet
	 * @param clazz
	 * @param rowIndex
	 * @return              返回当前写到第几行
	 * @throws Exception
	 */
	private int createHeader(Sheet sheet, Class<?> clazz, int rowIndex) throws Exception {
		Row row = sheet.createRow(rowIndex);
		Workbook workbook = sheet.getWorkbook();

		// 样式
		CellStyle cellStyle = null;
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);

		if (excelStyle==null) {
			cellStyle = new DefaultCellStyle().createHeaderStyle(workbook);
		} else {
			ICellStyle obj = (ICellStyle) Class.forName(excelStyle.cellStyle()).getDeclaredConstructor().newInstance();
			cellStyle = obj.createHeaderStyle(workbook);

			// 行高
			int height = excelStyle.headerHeight();
			if (height>0) {
				row.setHeight((short) (height * BASE_ROW_HEIGHT));
			}
		}

		skipMap.clear();
		int colIndex = 0;    // 列索引
		// 得到该类的所有成员变量
		Field[] fieldArr = clazz.getDeclaredFields();
		for (int i=0; i<fieldArr.length; i++) {
			Field field = fieldArr[i];

			// 设置类的私有属性可访问
			field.setAccessible(true);

			// 得到每一个成员变量上的 ExcelCell 注解
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}
			// 跳过被归入组的列
			if (skipMap.get(field.getName())!=null) {
				continue;
			}

			int sort = excelCell.sort()==0 ? colIndex : (excelCell.sort() - 1);

			// 设置单元格
			Cell cell = row.createCell(sort);
			cell.setCellValue(excelCell.name().length>0 ? excelCell.name()[0] : "");
			cell.setCellStyle(cellStyle);

			// 合并组
			int mergeCol = excelCell.group();
			if (mergeCol>1 && excelCell.sort()==0) {
				int num = 0;
				for (int j=(i+1); j<fieldArr.length; j++) {
					Field temp = fieldArr[j];
					temp.setAccessible(true);
					if (temp.getAnnotation(ExcelCell.class)==null) {
						continue;
					}

					skipMap.put(temp.getName(), temp.getName());

					num++;

					if (num==(mergeCol-1)) {
						break;
					}
				}
			}

			colIndex++;
		}

		return ++rowIndex;
	}

	/**
	 * 设置多行表头
	 * @param sheet
	 * @param clazz
	 * @param rowIndex
	 * @param headerRows    表头行数
	 * @return              返回当前写到第几行
	 * @throws Exception
	 */
	private int createHeaders(Sheet sheet, Class<?> clazz, int rowIndex, int headerRows) throws Exception {
		int rowIndexTemp = rowIndex;

		Workbook workbook = sheet.getWorkbook();

		// 行高
		int height = 0;
		// 样式
		CellStyle cellStyle = null;
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);

		if (excelStyle==null) {
			cellStyle = new DefaultCellStyle().createHeaderStyle(workbook);
		} else {
			ICellStyle obj = (ICellStyle) Class.forName(excelStyle.cellStyle()).getDeclaredConstructor().newInstance();
			cellStyle = obj.createHeaderStyle(workbook);

			// 行高
			height = excelStyle.headerHeight();
		}

		skipMap.clear();
		int colIndex = 0;    // 列索引
		// 得到该类的所有成员变量
		Field[] fieldArr = clazz.getDeclaredFields();
		for (int i=0; i<fieldArr.length; i++) {
			Field field = fieldArr[i];

			// 设置类的私有属性可访问
			field.setAccessible(true);

			// 得到每一个成员变量上的 ExcelCell 注解
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}
			// 跳过被归入组的列
			if (skipMap.get(field.getName())!=null) {
				continue;
			}

			int sort = excelCell.sort()==0 ? colIndex : (excelCell.sort() - 1);

			for (int n=0; n<headerRows; n++) {
				int rowIndexNew = rowIndexTemp + n;

				Row row = sheet.getRow(rowIndexNew);
				if (row==null) {
					row = sheet.createRow(rowIndexNew);

					// 行高
					if (height>0) {
						row.setHeight((short) (height * BASE_ROW_HEIGHT));
					}

					rowIndex++;
				}

				// 设置单元格
				String cellValue = n<excelCell.name().length ? excelCell.name()[n] : "";
				Cell cell = row.createCell(sort);
				cell.setCellValue(cellValue);
				cell.setCellStyle(cellStyle);
			}

			// 合并组
			int mergeCol = excelCell.group();
			if (mergeCol>1 && excelCell.sort()==0) {
				int num = 0;
				for (int j=(i+1); j<fieldArr.length; j++) {
					Field temp = fieldArr[j];
					temp.setAccessible(true);
					if (temp.getAnnotation(ExcelCell.class)==null) {
						continue;
					}

					skipMap.put(temp.getName(), temp.getName());

					num++;

					if (num==(mergeCol-1)) {
						break;
					}
				}
			}

			colIndex++;
		}

		// 设置表头合并
		SheetMergeHelper sheetMergeHelper = new SheetMergeHelper();
		sheetMergeHelper.setHeaderMerge(sheet, rowIndexTemp, headerRows);

		return rowIndex;
	}

	/**
	 * 设置数据
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param rowIndex
	 */
	@SuppressWarnings("unchecked")
	public void createData(Sheet sheet, Class<?> clazz, List<?> list, int rowIndex) throws Exception {
		this.createData(sheet, clazz, list, rowIndex, true);
	}

	/**
	 * 设置数据
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param rowIndex
	 * @param autoMergeRow 是否在本次数据写完后执行注解纵向合并
	 */
	@SuppressWarnings("unchecked")
	public void createData(Sheet sheet, Class<?> clazz, List<?> list, int rowIndex, boolean autoMergeRow) throws Exception {
		if (list==null || list.isEmpty()) {
			return;
		}

		// 行高
		int height = 0;
		// 注解导出面向普通列表和明细报表，默认按长文本自动撑高行高。
		boolean autoDataHeight = true;
		int maxDataHeight = 0;
		// 样式
		CellStyle cellStyle = null;
		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);
		if (excelStyle==null) {
			cellStyle = new DefaultCellStyle().createDataStyle(sheet.getWorkbook());
		} else {
			ICellStyle obj = (ICellStyle) Class.forName(excelStyle.cellStyle()).getDeclaredConstructor().newInstance();
			cellStyle = obj.createDataStyle(sheet.getWorkbook());

			// 行高
			height = excelStyle.dataHeight();
			autoDataHeight = excelStyle.autoDataHeight();
			maxDataHeight = excelStyle.maxDataHeight();
		}

		if (height>0) {
			autoDataHeight = false;
		}
		if (autoDataHeight) {
			cellStyle.setWrapText(true);
		}

		Field[] fieldArr = clazz.getDeclaredFields();

		CellHelper cellHelper = new CellHelper();

		Row row = null;           // 行
		Cell cell = null;         // 单元格
		int firstDataRow = rowIndex;
		int len = list.size();    // 数据行数
		for (int i=0; i<len; i++) {
			row = sheet.createRow(rowIndex);

			// 行高
			if (height>0) {
				row.setHeight((short) (height * BASE_ROW_HEIGHT));
			}
			int maxTextLineCount = MIN_AUTO_HEIGHT_LINES;

			Object entity = list.get(i);

			skipMap.clear();
			int colIndex = 0;    // 列索引
			for (int j=0; j<fieldArr.length; j++) {
				Field field = fieldArr[j];

				// 设置类的私有属性可访问
				field.setAccessible(true);

				// 得到每一个成员变量上的 ExcelCell 注解
				ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
				if (excelCell==null) {
					continue;
				}
				// 跳过被归入组的列
				if (skipMap.get(field.getName())!=null) {
					continue;
				}

				int sort = excelCell.sort()==0 ? colIndex : (excelCell.sort() - 1);

				// 创建单元格并设置值
				cell = row.createCell(sort);
				Object obj = this.convertExportValue(field, excelCell, field.get(entity));

				if (excelCell.type().contains("image")) {
					if (obj==null) {
						cell.setCellValue("");
					} else {
						String[] split = excelCell.type().split(",");
						Integer picWidth = split.length>1 && !"0".equals(split[1].trim()) ? Integer.valueOf(split[1].trim()) : null;
						Integer picHeight = split.length>2 && !"0".equals(split[2].trim()) ? Integer.valueOf(split[2].trim()) : null;
						if (obj instanceof List) {
							cellHelper.setImages(cell, (List<?>) obj, picWidth, picHeight);
						} else if (obj instanceof Picture) {
							Picture picture = (Picture) obj;
							if (picWidth!=null) {
								picture.setWidth(picWidth);
							}
							if (picHeight!=null) {
								picture.setHeight(picHeight);
							}
							cellHelper.setImage(cell, picture);
						} else {
							cellHelper.setImage(cell, String.valueOf(obj), picWidth, picHeight);
						}
					}
				} else {
					if (obj==null) {
						if (excelCell.defaultValue().length()>0) {
							cell.setCellValue(excelCell.defaultValue());   // 默认值
						} else {
							cell.setCellValue("");
						}
					} else if (obj instanceof String) {
						if ("".equals(obj) && excelCell.defaultValue().length()>0) {
							obj = excelCell.defaultValue();                // 默认值
						}

						cell.setCellValue((String) obj);
					}
					else if (obj instanceof Integer) {
						cell.setCellValue(Integer.parseInt(obj.toString()));
					}
					else if (obj instanceof Double) {
						cell.setCellValue(Double.parseDouble(obj.toString()));
					}
					else if (obj instanceof Long) {
						cell.setCellValue(Long.parseLong(obj.toString()));
					}
					else if (obj instanceof Float) {
						cell.setCellValue(Float.parseFloat(obj.toString()));
					}
					else if (obj instanceof BigDecimal) {
						cell.setCellValue(new BigDecimal(obj.toString()).doubleValue());
					}
					else if (obj instanceof LocalDateTime) {
						DateTimeFormatter dtf = (DateTimeFormatter) formatMap.get(String.valueOf(sort));
						cell.setCellValue(dtf==null ? obj.toString() : dtf.format((LocalDateTime) obj));
					}
					else if (obj instanceof LocalDate) {
						DateTimeFormatter dtf = (DateTimeFormatter) formatMap.get(String.valueOf(sort));
						cell.setCellValue(dtf==null ? obj.toString() : dtf.format((LocalDate) obj));
					}
					else if (obj instanceof Date) {
						SimpleDateFormat sdf = (SimpleDateFormat) formatMap.get(String.valueOf(sort));
						cell.setCellValue(sdf==null ? String.valueOf(((Date) obj).getTime()) : sdf.format((Date) obj));
					}
					else {
						cell.setCellValue(obj.toString());
					}

					// 值替换
					if (obj!=null && excelCell.replace().length>0) {
						Map<String, String> map = (Map<String, String>) replaceMap.get(String.valueOf(sort));
						if (map!=null && map.get(obj.toString())!=null) {
							cell.setCellValue(map.get(obj.toString()));
						}
					}
				}

				// 设置单元格样式
				cell.setCellStyle(cellStyle);

				int mergeCol = excelCell.group();
				if (mergeCol>1 && excelCell.sort()==0) {
					String mergeStr = obj==null ? "" : String.valueOf(obj);
					String separator = excelCell.separator();

					int num = 0;
					for (int k=(j+1); k<fieldArr.length; k++) {
						Field temp = fieldArr[k];
						temp.setAccessible(true);
						ExcelCell tempExcelCell = temp.getAnnotation(ExcelCell.class);
						if (tempExcelCell==null) {
							continue;
						}
						Object tempValue = this.convertExportValue(temp, tempExcelCell, temp.get(entity));
						String str = tempValue==null ? "" : String.valueOf(tempValue);
						mergeStr = mergeStr + separator + str;

						skipMap.put(temp.getName(), temp.getName());

						num++;

						if (num==(mergeCol-1)) {
							break;
						}
					}

					cell.setCellValue(mergeStr);
				}
				if (autoDataHeight && !excelCell.type().contains("image")) {
					maxTextLineCount = Math.max(maxTextLineCount, this.estimateCellLineCount(sheet, cell));
				}

				colIndex++;
			}
			if (autoDataHeight) {
				this.setAutoDataRowHeight(row, maxTextLineCount, maxDataHeight);
			}

			rowIndex++;
		}

		if (autoMergeRow) {
			this.mergeDataRows(sheet, clazz, firstDataRow, rowIndex - 1);
		}
	}

	/**
	 * 将字段原始值转换为导出值。
	 * 只有用户在 @ExcelCell.converter 中显式配置自定义转换器时才执行，默认转换器继续沿用原有写值逻辑。
	 * @param field 当前字段
	 * @param excelCell 字段注解
	 * @param fieldValue 字段原始值
	 * @return 写入Excel前的导出值
	 * @throws Exception
	 */
	private Object convertExportValue(Field field, ExcelCell excelCell, Object fieldValue) throws Exception {
		if (!this.hasCustomConverter(excelCell)) {
			return fieldValue;
		}

		// 自定义导出转换在基础类型写单元格前执行，转换器可以返回字符串、数字、日期或图片路径等现有导出逻辑支持的类型。
		ExcelValueConverter converter = excelCell.converter().getDeclaredConstructor().newInstance();
		return converter.convertToExcel(fieldValue, field, excelCell);
	}

	/**
	 * 判断字段是否配置了自定义转换器。
	 * 默认转换器只是占位，表示继续使用 officejj 内置的导入/导出转换逻辑。
	 * @param excelCell 字段注解
	 * @return 是否需要执行用户自定义转换器
	 */
	private boolean hasCustomConverter(ExcelCell excelCell) {
		return excelCell!=null && excelCell.converter()!=DefaultExcelValueConverter.class;
	}

	/**
	 * 根据注解配置自动调整列宽。
	 *     只在 @ExcelStyle(autoColumnWidth=true) 时生效，默认不影响旧的固定列宽逻辑。
	 *     这里扫描已经写入的表头和数据单元格，按最长显示文本估算列宽，并用 min/max 做上下限保护。
	 * @param sheet
	 * @param clazz
	 */
	public void applyAutoColumnWidth(Sheet sheet, Class<?> clazz) {
		if (sheet==null || clazz==null) {
			return;
		}

		ExcelStyle excelStyle = clazz.getAnnotation(ExcelStyle.class);
		if (!this.isAutoColumnWidth(excelStyle)) {
			return;
		}

		int minColumnWidth = this.resolveMinColumnWidth(excelStyle);
		int maxColumnWidth = this.resolveMaxColumnWidth(excelStyle);
		if (minColumnWidth>maxColumnWidth) {
			throw new IllegalArgumentException("minColumnWidth不能大于maxColumnWidth");
		}

		for (Integer colIndex : this.getExportColumnIndexList(clazz)) {
			int preferredWidth = this.calculatePreferredColumnWidth(sheet, colIndex, minColumnWidth, maxColumnWidth);
			sheet.setColumnWidth(colIndex, this.toColumnWidthUnits(preferredWidth));
		}
	}

	/**
	 * 取得注解导出的实际列索引。
	 *     这里复用导出时的 sort 和 group 规则，避免隐藏到组合并里的字段被重复计算。
	 * @param clazz
	 * @return
	 */
	private List<Integer> getExportColumnIndexList(Class<?> clazz) {
		List<Integer> colIndexList = new ArrayList<Integer>();
		Map<String, String> groupSkipMap = new HashMap<String, String>();

		int colIndex = 0;
		Field[] fieldArr = clazz.getDeclaredFields();
		for (int i=0; i<fieldArr.length; i++) {
			Field field = fieldArr[i];
			field.setAccessible(true);

			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}
			if (groupSkipMap.get(field.getName())!=null) {
				continue;
			}

			int sort = excelCell.sort()==0 ? colIndex : (excelCell.sort() - 1);
			colIndexList.add(sort);

			int mergeCol = excelCell.group();
			if (mergeCol>1 && excelCell.sort()==0) {
				int num = 0;
				for (int j=(i+1); j<fieldArr.length; j++) {
					Field temp = fieldArr[j];
					temp.setAccessible(true);
					if (temp.getAnnotation(ExcelCell.class)==null) {
						continue;
					}

					groupSkipMap.put(temp.getName(), temp.getName());
					num++;
					if (num==(mergeCol-1)) {
						break;
					}
				}
			}

			colIndex++;
		}

		return colIndexList;
	}

	/**
	 * 计算单列的推荐宽度。
	 *     标题行通常是横向合并单元格，直接参与会把第一列错误撑宽，所以横向合并内容不参与列宽估算。
	 * @param sheet
	 * @param colIndex
	 * @param minColumnWidth
	 * @param maxColumnWidth
	 * @return
	 */
	private int calculatePreferredColumnWidth(Sheet sheet, int colIndex, int minColumnWidth, int maxColumnWidth) {
		double maxTextWidth = 0.0D;
		int lastRowNum = sheet.getLastRowNum();
		for (int rowIndex=0; rowIndex<=lastRowNum; rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (row==null) {
				continue;
			}

			Cell cell = row.getCell(colIndex);
			if (cell==null || this.isHorizontalMergedCell(sheet, cell)) {
				continue;
			}

			maxTextWidth = Math.max(maxTextWidth, this.estimateSingleLineTextWidth(this.getCellDisplayText(cell)));
		}

		int preferredWidth = (int) Math.ceil(maxTextWidth + AUTO_COLUMN_WIDTH_PADDING_CHAR_WIDTH);
		preferredWidth = Math.max(minColumnWidth, preferredWidth);
		preferredWidth = Math.min(maxColumnWidth, preferredWidth);

		return preferredWidth;
	}

	/**
	 * 判断当前单元格是否处在横向合并区域内。
	 * @param sheet
	 * @param cell
	 * @return
	 */
	private boolean isHorizontalMergedCell(Sheet sheet, Cell cell) {
		CellRangeAddress region = this.getMergedRegion(sheet, cell.getRowIndex(), cell.getColumnIndex());
		return region!=null && region.getFirstColumn()!=region.getLastColumn();
	}

	/**
	 * 按单行展示估算文本宽度。
	 *     不自动撑高行高时，显式换行文本按最长一行计算列宽，避免把多行内容宽度累加得过大。
	 * @param text
	 * @return
	 */
	private double estimateSingleLineTextWidth(String text) {
		if (text==null || text.length()==0) {
			return 0.0D;
		}

		double maxLineWidth = 0.0D;
		String[] lines = text.split("\\r\\n|\\n|\\r", -1);
		for (String line : lines) {
			double lineWidth = 0.0D;
			for (int i=0; i<line.length(); i++) {
				lineWidth += this.getCharWidth(line.charAt(i));
			}
			maxLineWidth = Math.max(maxLineWidth, lineWidth);
		}

		return maxLineWidth;
	}

	/**
	 * 判断是否启用注解自动列宽。
	 * @param excelStyle
	 * @return
	 */
	private boolean isAutoColumnWidth(ExcelStyle excelStyle) {
		return excelStyle!=null && excelStyle.autoColumnWidth();
	}

	/**
	 * 解析自动列宽最小值。
	 * @param excelStyle
	 * @return
	 */
	private int resolveMinColumnWidth(ExcelStyle excelStyle) {
		if (excelStyle==null || excelStyle.minColumnWidth()<=0) {
			return 1;
		}

		return Math.min(excelStyle.minColumnWidth(), MAX_EXCEL_COLUMN_WIDTH);
	}

	/**
	 * 解析自动列宽最大值。
	 * @param excelStyle
	 * @return
	 */
	private int resolveMaxColumnWidth(ExcelStyle excelStyle) {
		if (excelStyle==null || excelStyle.maxColumnWidth()<=0) {
			return MAX_EXCEL_COLUMN_WIDTH;
		}

		return Math.min(excelStyle.maxColumnWidth(), MAX_EXCEL_COLUMN_WIDTH);
	}

	/**
	 * 转换为 POI setColumnWidth 需要的 1/256 字符宽度单位。
	 * @param columnWidth
	 * @return
	 */
	private int toColumnWidthUnits(int columnWidth) {
		int safeColumnWidth = Math.max(1, Math.min(columnWidth, MAX_EXCEL_COLUMN_WIDTH));
		return safeColumnWidth * BASE_COLUMN_WIDTH;
	}

	/**
	 * 根据单元格文本和列宽估算需要显示几行。
	 * POI 不会像 Excel 客户端那样自动计算行高，因此导出阶段需要做一个保守估算。
	 * @param sheet
	 * @param cell
	 * @return
	 */
	private int estimateCellLineCount(Sheet sheet, Cell cell) {
		String text = this.getCellDisplayText(cell);
		if (text==null || text.length()==0) {
			return MIN_AUTO_HEIGHT_LINES;
		}

		double columnWidth = Math.max(1.0D, this.getCellWidthInChars(sheet, cell) - CELL_PADDING_CHAR_WIDTH);
		String[] lines = text.split("\\r\\n|\\n|\\r", -1);
		int lineCount = 0;
		for (String line : lines) {
			lineCount += this.estimateWrappedLineCount(line, columnWidth);
		}

		return Math.max(MIN_AUTO_HEIGHT_LINES, lineCount);
	}

	/**
	 * 得到单元格最终显示文本。
	 * @param cell
	 * @return
	 */
	private String getCellDisplayText(Cell cell) {
		if (cell==null) {
			return "";
		}
		CellType cellType = cell.getCellType();
		if (cellType==CellType.STRING) {
			return cell.getStringCellValue();
		}
		if (cellType==CellType.NUMERIC) {
			return String.valueOf(cell.getNumericCellValue());
		}
		if (cellType==CellType.BOOLEAN) {
			return String.valueOf(cell.getBooleanCellValue());
		}
		if (cellType==CellType.FORMULA) {
			return cell.getCellFormula();
		}

		return "";
	}

	/**
	 * 得到单元格可用列宽，合并单元格按合并区域总宽度估算。
	 * @param sheet
	 * @param cell
	 * @return
	 */
	private double getCellWidthInChars(Sheet sheet, Cell cell) {
		CellRangeAddress region = this.getMergedRegion(sheet, cell.getRowIndex(), cell.getColumnIndex());
		if (region==null) {
			return sheet.getColumnWidth(cell.getColumnIndex()) / 256.0D;
		}

		double width = 0.0D;
		for (int i=region.getFirstColumn(); i<=region.getLastColumn(); i++) {
			width += sheet.getColumnWidth(i) / 256.0D;
		}

		return width;
	}

	/**
	 * 查找单元格所在的合并区域。
	 * @param sheet
	 * @param rowIndex
	 * @param colIndex
	 * @return
	 */
	private CellRangeAddress getMergedRegion(Sheet sheet, int rowIndex, int colIndex) {
		for (CellRangeAddress region : sheet.getMergedRegions()) {
			if (region.isInRange(rowIndex, colIndex)) {
				return region;
			}
		}

		return null;
	}

	/**
	 * 按列宽估算一段无显式换行文本会被折成几行。
	 * @param text
	 * @param columnWidth
	 * @return
	 */
	private int estimateWrappedLineCount(String text, double columnWidth) {
		if (text==null || text.length()==0) {
			return 1;
		}

		int lineCount = 1;
		double currentWidth = 0.0D;
		for (int i=0; i<text.length(); i++) {
			double charWidth = this.getCharWidth(text.charAt(i));
			if (currentWidth>0 && currentWidth + charWidth > columnWidth) {
				lineCount++;
				currentWidth = charWidth;
			} else {
				currentWidth += charWidth;
			}
		}

		return lineCount;
	}

	/**
	 * 粗略估算字符宽度，中文/日文/韩文等全角字符按2个英文字符计算。
	 * @param ch
	 * @return
	 */
	private double getCharWidth(char ch) {
		if (ch=='\t') {
			return 4.0D;
		}
		Character.UnicodeScript script = Character.UnicodeScript.of(ch);
		if (script==Character.UnicodeScript.HAN
				|| script==Character.UnicodeScript.HIRAGANA
				|| script==Character.UnicodeScript.KATAKANA
				|| script==Character.UnicodeScript.HANGUL) {
			return 2.0D;
		}
		if (ch>255) {
			return 1.5D;
		}

		return 1.0D;
	}

	/**
	 * 设置自动数据行高，保留图片等逻辑已经撑开的高度，不会把行高调小。
	 * @param row
	 * @param lineCount
	 * @param maxDataHeight 最大高度，单位：磅，0表示不限制
	 */
	private void setAutoDataRowHeight(Row row, int lineCount, int maxDataHeight) {
		float baseHeight = row.getSheet().getDefaultRowHeightInPoints();
		if (baseHeight<=0) {
			baseHeight = 15.0F;
		}

		float targetHeight = baseHeight * Math.max(MIN_AUTO_HEIGHT_LINES, lineCount);
		if (maxDataHeight>0) {
			targetHeight = Math.min(targetHeight, maxDataHeight);
		}

		if (targetHeight>row.getHeightInPoints()) {
			row.setHeightInPoints(targetHeight);
		}
	}

	/**
	 * 按注解配置自动纵向合并数据行。
	 * 只合并当前写入的数据区域，不需要调用方再手动计算截止行。
	 * @param sheet
	 * @param clazz
	 * @param firstRow 起始行（从0开始计算）
	 * @param lastRow  终止行（从0开始计算）
	 */
	public void mergeDataRows(Sheet sheet, Class<?> clazz, int firstRow, int lastRow) {
		if (sheet==null || clazz==null || firstRow>=lastRow) {
			return;
		}

		List<MergeRowColumn> mergeColumnList = this.getMergeRowColumnList(clazz);
		if (mergeColumnList.isEmpty()) {
			return;
		}

		SheetMergeHelper sheetMergeHelper = new SheetMergeHelper();
		for (MergeRowColumn mergeRowColumn : mergeColumnList) {
			sheetMergeHelper.setAutoMergeCol(sheet, mergeRowColumn.getColIndex(), firstRow, lastRow, null, mergeRowColumn.getMergeByColIndexArr());
		}
	}

	/**
	 * 解析需要纵向合并的列。
	 * 这里复用导出列的排序规则，保证注解里的 mergeBy={1} 指向最终Excel中的第1列。
	 * @param clazz
	 * @return
	 */
	private List<MergeRowColumn> getMergeRowColumnList(Class<?> clazz) {
		List<MergeRowColumn> list = new ArrayList<MergeRowColumn>();
		Map<String, String> groupSkipMap = new HashMap<String, String>();

		int colIndex = 0;
		Field[] fieldArr = clazz.getDeclaredFields();
		for (int i=0; i<fieldArr.length; i++) {
			Field field = fieldArr[i];
			field.setAccessible(true);

			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}
			if (groupSkipMap.get(field.getName())!=null) {
				continue;
			}

			int sort = excelCell.sort()==0 ? colIndex : (excelCell.sort() - 1);
			if (excelCell.mergeRow()) {
				list.add(new MergeRowColumn(sort, this.toZeroBasedMergeBy(excelCell.mergeBy())));
			}

			// 横向组合并会隐藏后续字段，这些字段不再参与最终列序号计算。
			int mergeCol = excelCell.group();
			if (mergeCol>1 && excelCell.sort()==0) {
				int num = 0;
				for (int j=(i+1); j<fieldArr.length; j++) {
					Field temp = fieldArr[j];
					temp.setAccessible(true);
					if (temp.getAnnotation(ExcelCell.class)==null) {
						continue;
					}

					groupSkipMap.put(temp.getName(), temp.getName());
					num++;
					if (num==(mergeCol-1)) {
						break;
					}
				}
			}

			colIndex++;
		}

		return list;
	}

	/**
	 * 把注解中面向用户的1-based依赖列号转换为内部0-based列号。
	 * @param mergeBy
	 * @return
	 */
	private int[] toZeroBasedMergeBy(int[] mergeBy) {
		if (mergeBy==null || mergeBy.length==0) {
			return new int[0];
		}

		int[] result = new int[mergeBy.length];
		for (int i=0; i<mergeBy.length; i++) {
			if (mergeBy[i]<1) {
				throw new IllegalArgumentException("mergeBy列号必须从1开始：" + mergeBy[i]);
			}
			result[i] = mergeBy[i] - 1;
		}

		return result;
	}

	/**
	 * 纵向合并列配置。
	 */
	private static class MergeRowColumn {
		private final int colIndex;
		private final int[] mergeByColIndexArr;

		private MergeRowColumn(int colIndex, int[] mergeByColIndexArr) {
			this.colIndex = colIndex;
			this.mergeByColIndexArr = mergeByColIndexArr;
		}

		private int getColIndex() {
			return colIndex;
		}

		private int[] getMergeByColIndexArr() {
			return mergeByColIndexArr;
		}
	}
}
