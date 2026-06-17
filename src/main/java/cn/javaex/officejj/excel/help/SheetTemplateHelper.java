package cn.javaex.officejj.excel.help;

import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetView;

import cn.javaex.officejj.common.entity.Font;
import cn.javaex.officejj.common.entity.Picture;
import cn.javaex.officejj.common.util.PropertyHandler;
import cn.javaex.officejj.excel.ExcelUtils;
import cn.javaex.officejj.excel.entity.SheetPrintArea;

/**
 * 模板替换写入Excel
 *
 * @author 陈霓清
 */
public class SheetTemplateHelper extends SheetHelper {
	/** AWT字体测量上下文，用于按真实字体宽度估算Excel换行。 */
	private static final java.awt.font.FontRenderContext FONT_RENDER_CONTEXT = new java.awt.font.FontRenderContext(null, true, true);
	/** 模板自动行高最小行数 */
	private static final int MIN_AUTO_HEIGHT_LINES = 1;
	/** 单元格左右内边距，按像素扣减，避免文本贴边时估算偏小。 */
	private static final double CELL_PADDING_PIXEL_WIDTH = 8.0D;
	/** 模板打印视图里的中文换行常比标准列宽更早发生，字符宽度估算按更窄的有效宽度处理。 */
	private static final double TEMPLATE_CHAR_WIDTH_FACTOR = 0.60D;
	/** 单元格左右内边距折算的字符数，避免文本贴边时估算偏小。 */
	private static final double CELL_PADDING_CHAR_WIDTH = 3.0D;
	/** 长文本兜底：每12个字符至少按一行估算，优先避免任何截断。 */
	private static final double LONG_TEXT_CHARS_PER_LINE = 12.0D;
	/** Excel单行最大行高约409磅，超过后写入会不可靠，必须截到上限。 */
	private static final float MAX_EXCEL_ROW_HEIGHT_POINTS = 409.0F;
	/** AWT 72DPI用户空间转POI列宽像素的系数。 */
	private static final double POINTS_TO_PIXELS = 96.0D / 72.0D;
	/** 行高额外上下边距，避免贴边和少量字体渲染差异。 */
	private static final double ROW_HEIGHT_PADDING_POINTS = 4.0D;

	/**
	 * 替换占位符（相同的占位符，只处理一次）
	 */
	@Override
	public void write(Sheet sheet, Map<String, Object> param) {
		this.write(sheet, param, false);
	}

	/**
	 * 替换占位符（相同的占位符，只处理一次）。
	 * 自动行高默认关闭，避免模板打印区域被长文本撑破；调用方显式传 true 时才处理。
	 */
	@Override
	public void write(Sheet sheet, Map<String, Object> param, boolean autoDataHeight) {
		Set<String> handledListKeys = new HashSet<>();
		Set<String> handledTextKeys = new HashSet<>();
		Map<Short, CellStyle> wrapStyleMap = new HashMap<Short, CellStyle>();

		CellHelper cellHelper = new CellHelper();

		Map<String, List<TemplateListCell>> listMap = new LinkedHashMap<String, List<TemplateListCell>>();

		Row row = null;
		Cell cell = null;
		int index = 0;
		int lastRowNum = sheet.getLastRowNum();

		while (index <= lastRowNum) {
			row = sheet.getRow(index++);
			if (row==null) {
				continue;
			}

			List<TemplateListCell> list = new ArrayList<TemplateListCell>();
			String tempListKey = "";
			String listKey = "";

			int startCol = row.getFirstCellNum();    // 索引
			int endCol = row.getLastCellNum();       // 从1开始计算
			if (startCol<0 || endCol<0) {
				continue;
			}
			for (int i=startCol; i<endCol; i++) {
				if (row.getCell(i)==null) {
					continue;
				}

				// 得到单元格的内容
				String cellValue = this.getTemplateCellText(row.getCell(i));

				// 如果单元格的内容不包含 ${xxx}，则跳过
				if (!(cellValue.contains("${") && cellValue.contains("}"))) {
					continue;
				}

				// 获取该单元格内的所有占位符变量
				List<String> placeholders = cellHelper.getPlaceholders(cellValue);
				if (placeholders.isEmpty()) {
					continue;
				}

				// 只有根对象真实为 List 时，${list.name} 才表示列表遍历。
				// 普通对象属性路径如 ${project.className} 应走直接替换，不能强转 List。
				String currentListKey = this.getListKey(placeholders, param);
				if (currentListKey!=null) {
					listKey = currentListKey;

					// 如果该key已经处理过（已入listMap并执行过 setListValue），后续重复区域直接忽略
					if (handledListKeys.contains(listKey)) {
						continue;
					}

					if (!"".equals(tempListKey) && !"".equals(listKey) && !tempListKey.equals(listKey)) {
						// flush 上一个key
						listMap.put(tempListKey, list);
						handledListKeys.add(tempListKey); // 标记为已处理

						this.setListValue(sheet, listMap, param, wrapStyleMap, autoDataHeight);
						listMap.clear();

						tempListKey = "";
						list = new ArrayList<TemplateListCell>();
					} else {
						tempListKey = listKey;
					}

					list.add(new TemplateListCell(listKey, row.getRowNum(), i, cellValue, placeholders));
				}
				// 直接替换（非list）
				else {
					// 独占一格：一定只有一个 key 时
					if (cellValue.equals("${" + placeholders.get(0) + "}")) {
						String key = placeholders.get(0);
						if (handledTextKeys.contains(key)) {
							continue;
						}
						cell = sheet.getRow(row.getRowNum()).getCell(i);
						Object value = PropertyHandler.getValue(param, key);
						cellHelper.setValue(cell, value);
						if (autoDataHeight) {
							this.autoFitTemplateCell(cell, value, wrapStyleMap);
						}
						handledTextKeys.add(key); // 标记为已处理
					}
					// 非独占：可能有多个占位符
					else {
						// 只替换“尚未处理过”的占位符
						List<String> toReplace = new ArrayList<>();
						for (String key : placeholders) {
							if (!handledTextKeys.contains(key)) {
								toReplace.add(key);
							}
						}
						if (toReplace.isEmpty()) {
							continue;
						}

						cell = sheet.getRow(row.getRowNum()).getCell(i);
						cellHelper.setValue(cell, toReplace, param);
						if (autoDataHeight) {
							this.autoFitTemplateCell(cell, null, wrapStyleMap);
						}

						// 替换后把这些 key 标记为已处理
						handledTextKeys.addAll(toReplace);
					}
				}
			}

			if (!"".equals(listKey) && !handledListKeys.contains(listKey)) {
				listMap.put(listKey, list);
				handledListKeys.add(listKey);
			}
		}

		if (listMap.isEmpty()==false) {
			this.setListValue(sheet, listMap, param, wrapStyleMap, autoDataHeight);
			listMap.clear();
		}
		if (autoDataHeight) {
			adjustPrintScaleForManualPageBreaks(sheet);
		}
	}

	/**
	 * 替换模板中的占位符（list遍历）
	 * @param sheet
	 * @param listMap
	 * @param param
	 */
	@SuppressWarnings("unchecked")
	private void setListValue(Sheet sheet, Map<String, List<TemplateListCell>> listMap, Map<String, Object> param, Map<Short, CellStyle> wrapStyleMap, boolean autoDataHeight) {
		CellHelper cellHelper = new CellHelper();

		// LinkedHashMap倒序遍历
		ListIterator<Map.Entry<String, List<TemplateListCell>>> iterator = new ArrayList<Map.Entry<String, List<TemplateListCell>>>(listMap.entrySet()).listIterator(listMap.size());
		while (iterator.hasPrevious()) {
			Map.Entry<String, List<TemplateListCell>> entry = iterator.previous();

			// 1.0 取出需要遍历的list数据
			Object listObj = PropertyHandler.getValue(param, entry.getKey());
			if (!(listObj instanceof List)) {
				continue;
			}
			List<Map<String, Object>> list = (List<Map<String, Object>>) listObj;
			if (list==null || list.isEmpty()) {
				continue;
			}

			// 2.0 遍历取出每一条数据并设置值
			int len = list.size();
			List<TemplateListCell> placeholders = entry.getValue();
			TemplateListLayout layout = this.getTemplateListLayout(sheet, placeholders);
			for (int i=0; i<len; i++) {
				Map<String, Object> dataMap = this.convertToMap(list.get(i));
				int itemStartRow = layout.getStartRow() + i * layout.getBlockRows();

				for (TemplateListCell templateListCell : placeholders) {
					int rowIndex = itemStartRow + templateListCell.getRowIndex() - layout.getStartRow();
					int colIndex = templateListCell.getColIndex();
					Row row = sheet.getRow(rowIndex);
					if (row == null) row = sheet.createRow(rowIndex);
					Cell cell = row.getCell(colIndex);
					if (cell == null) cell = row.createCell(colIndex);

					Object value = this.getListCellValue(templateListCell, dataMap, param);
					cellHelper.setValue(cell, value);
					if (autoDataHeight) {
						this.applyWrapText(cell, wrapStyleMap);
					}
				}
			}
			if (autoDataHeight) {
				this.autoFitTemplateListRows(sheet, placeholders, len, wrapStyleMap);
			}
		}
	}

	/**
	 * 判断占位符是否为列表遍历占位符。
	 * 带点号的占位符同时可能表示对象属性路径，必须以参数实际类型为准。
	 * @param placeholder 占位符内容，不包含 ${}
	 * @param param 模板参数
	 * @return 是否需要按列表遍历处理
	 */
	private boolean isListPlaceholder(String placeholder, Map<String, Object> param) {
		if (placeholder==null || !placeholder.contains(".")) {
			return false;
		}
		String listKey = placeholder.split("\\.", 2)[0];
		Object listObj = PropertyHandler.getValue(param, listKey);
		return listObj instanceof List;
	}

	/**
	 * 从单元格占位符中找出列表参数名。
	 * @param placeholders 单元格内的占位符
	 * @param param 模板参数
	 * @return 列表参数名，不是列表占位符时返回null
	 */
	private String getListKey(List<String> placeholders, Map<String, Object> param) {
		for (String placeholder : placeholders) {
			if (this.isListPlaceholder(placeholder, param)) {
				return placeholder.split("\\.", 2)[0];
			}
		}

		return null;
	}

	/**
	 * 获取模板单元格原始文本，保留共享单元格中的固定文案。
	 * @param cell 模板单元格
	 * @return 模板文本
	 */
	private String getTemplateCellText(Cell cell) {
		if (cell==null) {
			return "";
		}
		if (cell.getCellType()==CellType.STRING) {
			return cell.getStringCellValue();
		}

		return ExcelUtils.getCellValue(cell);
	}

	/**
	 * 生成列表模板单元格的写入值。
	 * 独占占位符保留原对象类型，图片、数字、日期等仍由 CellHelper 写入；共享单元格统一替换成最终文本。
	 * @param templateListCell 列表模板单元格
	 * @param dataMap 当前列表行数据
	 * @param param 模板参数
	 * @return 写入单元格的值
	 */
	private Object getListCellValue(TemplateListCell templateListCell, Map<String, Object> dataMap, Map<String, Object> param) {
		if (templateListCell.isExclusive()) {
			return PropertyHandler.getValue(dataMap, templateListCell.getExclusiveAttributeKey());
		}

		String cellValue = templateListCell.getCellTemplate();
		for (String placeholder : templateListCell.getPlaceholders()) {
			Object value = this.getPlaceholderValue(templateListCell.getListKey(), placeholder, dataMap, param);
			cellValue = cellValue.replace("${" + placeholder + "}", this.toPlaceholderText(value));
		}

		return cellValue;
	}

	/**
	 * 根据占位符来源获取替换值。
	 * @param listKey 当前列表参数名
	 * @param placeholder 占位符
	 * @param dataMap 当前列表行数据
	 * @param param 模板参数
	 * @return 替换值
	 */
	private Object getPlaceholderValue(String listKey, String placeholder, Map<String, Object> dataMap, Map<String, Object> param) {
		String prefix = listKey + ".";
		if (placeholder.startsWith(prefix)) {
			return PropertyHandler.getValue(dataMap, placeholder.substring(prefix.length()));
		}

		return PropertyHandler.getValue(param, placeholder);
	}

	/**
	 * 共享单元格只能写成文本，空值按空字符串替换。
	 * @param value 占位符值
	 * @return 占位符文本
	 */
	private String toPlaceholderText(Object value) {
		if (value==null) {
			return "";
		}
		if (value instanceof Font) {
			return ((Font) value).getText();
		}

		return value.toString();
	}

	/**
	 * 模板写值后自动适配单元格行高。
	 * @param cell
	 * @param value
	 * @param wrapStyleMap
	 */
	private void autoFitTemplateCell(Cell cell, Object value, Map<Short, CellStyle> wrapStyleMap) {
		this.autoFitTemplateCell(cell, value, 1, wrapStyleMap);
	}

	/**
	 * 模板写值后自动适配单元格行高。
	 * 图片由图片写入逻辑负责撑开行高，这里只处理文本。
	 * @param cell
	 * @param value
	 * @param rowSpan
	 * @param wrapStyleMap
	 */
	private void autoFitTemplateCell(Cell cell, Object value, int rowSpan, Map<Short, CellStyle> wrapStyleMap) {
		if (cell==null || this.isImageValue(value)) {
			return;
		}

		String text = this.getCellDisplayText(cell);
		if (text==null || text.length()==0) {
			return;
		}

		this.applyWrapText(cell, wrapStyleMap);
		float targetHeight = this.estimateCellHeightInPoints(cell.getSheet(), cell, text);
		this.setAutoTemplateRowHeight(cell.getSheet(), cell.getRowIndex(), Math.max(1, rowSpan), targetHeight);
	}

	/**
	 * list模板写值完成后，按本次写入的数据行重新计算行高。
	 * 同一行存在多个长文本列时，扫描整行文本并取最大需求，避免某些列未进入占位符集合导致行高偏小。
	 * @param sheet
	 * @param templateListCells
	 * @param dataSize
	 */
	private void autoFitTemplateListRows(Sheet sheet, List<TemplateListCell> templateListCells, int dataSize, Map<Short, CellStyle> wrapStyleMap) {
		TemplateListLayout layout = this.getTemplateListLayout(sheet, templateListCells);
		for (int i=0; i<dataSize; i++) {
			int blockStartRow = layout.getStartRow() + i * layout.getBlockRows();
			this.autoFitTemplateListBlockRows(sheet, blockStartRow, layout.getBlockRows(), wrapStyleMap);
		}
	}

	/**
	 * 按单条list数据对应的模板块扫描文本并调整行高。
	 * TIS这类模板可能一条数据跨多行或包含纵向合并单元格，必须按整块扫描并把高度分摊到合并区域覆盖的所有行。
	 * @param sheet
	 * @param blockStartRow
	 * @param blockRows
	 * @param wrapStyleMap
	 */
	private void autoFitTemplateListBlockRows(Sheet sheet, int blockStartRow, int blockRows, Map<Short, CellStyle> wrapStyleMap) {
		int blockEndRow = blockStartRow + Math.max(1, blockRows) - 1;
		for (int rowIndex=blockStartRow; rowIndex<=blockEndRow; rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (row==null) {
				continue;
			}
			int firstCellNum = row.getFirstCellNum();
			int lastCellNum = row.getLastCellNum();
			if (firstCellNum<0 || lastCellNum<0) {
				continue;
			}
			for (int colIndex=firstCellNum; colIndex<lastCellNum; colIndex++) {
				Cell cell = row.getCell(colIndex);
				if (cell==null) {
					continue;
				}

				CellRangeAddress mergedRegion = this.getMergedRegion(sheet, rowIndex, colIndex);
				if (mergedRegion!=null && (mergedRegion.getFirstRow()!=rowIndex || mergedRegion.getFirstColumn()!=colIndex)) {
					continue;
				}

				String text = this.getCellDisplayText(cell);
				if (text==null || text.length()==0) {
					continue;
				}

				this.applyWrapText(cell, wrapStyleMap);
				float targetHeight = this.estimateCellHeightInPoints(sheet, cell, text);
				int firstRow = mergedRegion==null ? rowIndex : mergedRegion.getFirstRow();
				int rowSpan = mergedRegion==null ? 1 : mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
				this.setAutoTemplateRowHeight(sheet, firstRow, rowSpan, targetHeight);
			}
		}
	}

	/**
	 * 判断当前值是否是图片，图片高度由图片写入逻辑控制。
	 * @param value
	 * @return
	 */
	private boolean isImageValue(Object value) {
		if (value instanceof Picture || value instanceof byte[]) {
			return true;
		}
		if (value instanceof List) {
			List<?> list = (List<?>) value;
			if (list.isEmpty()) {
				return false;
			}
			for (Object item : list) {
				if (!(item instanceof Picture) && !(item instanceof byte[])) {
					return false;
				}
			}
			return true;
		}

		return false;
	}

	/**
	 * 给模板单元格开启自动换行，保留原有样式。
	 * @param cell
	 * @param wrapStyleMap
	 */
	private void applyWrapText(Cell cell, Map<Short, CellStyle> wrapStyleMap) {
		CellStyle cellStyle = cell.getCellStyle();
		if (cellStyle==null) {
			return;
		}
		if (cellStyle.getWrapText() && !cellStyle.getShrinkToFit()) {
			return;
		}

		CellStyle wrapStyle = wrapStyleMap.get(cellStyle.getIndex());
		if (wrapStyle==null) {
			wrapStyle = cell.getSheet().getWorkbook().createCellStyle();
			wrapStyle.cloneStyleFrom(cellStyle);
			wrapStyle.setWrapText(true);
			wrapStyle.setShrinkToFit(false);
			wrapStyleMap.put(cellStyle.getIndex(), wrapStyle);
		}
		cell.setCellStyle(wrapStyle);
	}

	/**
	 * 根据单元格文本和列宽估算需要显示几行。
	 * @param sheet
	 * @param cell
	 * @param text
	 * @return
	 */
	private float estimateCellHeightInPoints(Sheet sheet, Cell cell, String text) {
		float defaultHeight = sheet.getDefaultRowHeightInPoints();
		if (defaultHeight<=0) {
			defaultHeight = 15.0F;
		}
		if (text==null || text.length()==0) {
			return defaultHeight;
		}

		double availableWidthPx = Math.max(1.0D, this.getCellWidthInPixels(sheet, cell) - CELL_PADDING_PIXEL_WIDTH);
		java.awt.Font awtFont = this.getAwtFont(sheet, cell);
		String[] lines = text.split("\\r\\n|\\n|\\r", -1);

		int fontMeasuredLineCount = this.calculateTotalLines(text, awtFont, availableWidthPx);
		int charLineCount = this.estimateCellLineCountByChars(sheet, cell, lines);
		int textLengthLineCount = this.estimateCellLineCountByTextLength(sheet, cell, lines);
		int totalLines = Math.max(MIN_AUTO_HEIGHT_LINES, Math.max(Math.max(fontMeasuredLineCount, charLineCount), textLengthLineCount));

		double correctedLineHeight = this.getCorrectedLineHeight(sheet, awtFont);
		double targetHeight = totalLines * correctedLineHeight + ROW_HEIGHT_PADDING_POINTS;
		double roundedHeight = Math.ceil(targetHeight * 20.0D) / 20.0D;
		return (float) Math.min(Math.max(defaultHeight, roundedHeight), MAX_EXCEL_ROW_HEIGHT_POINTS);
	}

	private int calculateTotalLines(String text, java.awt.Font awtFont, double availableWidthPx) {
		String[] lines = text.split("\\r\\n|\\n|\\r", -1);
		int totalLines = 0;
		for (String line : lines) {
			if (line==null || line.length()==0) {
				totalLines++;
				continue;
			}

			java.awt.geom.Rectangle2D bounds = awtFont.getStringBounds(line, FONT_RENDER_CONTEXT);
			double textWidthPx = bounds.getWidth() * POINTS_TO_PIXELS;
			int continuousLines = (int) Math.ceil(textWidthPx / availableWidthPx);

			int discreteLines = continuousLines;
			int charCount = line.length();
			if (charCount>0 && textWidthPx>0) {
				double avgCharWidthPx = textWidthPx / charCount;
				int charsPerLine = Math.max(1, (int) Math.floor(availableWidthPx / avgCharWidthPx));
				discreteLines = (int) Math.ceil((double) charCount / charsPerLine);
			}

			totalLines += Math.max(1, Math.max(continuousLines, discreteLines));
		}

		return totalLines;
	}

	private double getCorrectedLineHeight(Sheet sheet, java.awt.Font awtFont) {
		float defaultHeight = sheet.getDefaultRowHeightInPoints();
		if (defaultHeight<=0) {
			defaultHeight = 15.0F;
		}
		java.awt.Font normalFont = this.getWorkbookDefaultAwtFont(sheet);
		java.awt.font.LineMetrics normalLineMetrics = normalFont.getLineMetrics("Ay", FONT_RENDER_CONTEXT);
		double normalAwtLineHeight = normalLineMetrics.getHeight();
		if (normalAwtLineHeight<=0) {
			return defaultHeight;
		}

		java.awt.font.LineMetrics lineMetrics = awtFont.getLineMetrics("Ay", FONT_RENDER_CONTEXT);
		double awtLineHeight = lineMetrics.getHeight();
		if (awtLineHeight<=0) {
			return defaultHeight;
		}

		// RowHeightCalculator里的校正系数应基于工作簿默认字体计算，再应用到当前单元格字体。
		// 不能用当前字体自己除自己，否则实际字体差异会被抵消，宋体/加粗等多行文本容易少算行高。
		double correction = defaultHeight / normalAwtLineHeight;
		return Math.max(defaultHeight, awtLineHeight * correction);
	}

	/**
	 * 长文本最低行数兜底。
	 * 有些模板字体、打印视图或缩放设置会让真实可显示字符数明显少于列宽估算值，这里直接按字数给保守下限。
	 * @param lines
	 * @return
	 */
	private int estimateCellLineCountByTextLength(Sheet sheet, Cell cell, String[] lines) {
		double charsPerLine = this.getConservativeCharsPerLine(sheet, cell);
		int lineCount = 0;
		for (String line : lines) {
			if (line==null || line.length()==0) {
				lineCount++;
			} else {
				lineCount += Math.max(1, (int) Math.ceil(line.length() / charsPerLine));
			}
		}

		return lineCount;
	}

	/**
	 * 按列宽给长文本兜底估算每行字符数。
	 * 窄列需要显著减少每行字符数，避免后面的窄列长文本被前面宽列的行高覆盖。
	 * @param sheet
	 * @param cell
	 * @return
	 */
	private double getConservativeCharsPerLine(Sheet sheet, Cell cell) {
		double columnWidth = Math.max(1.0D, this.getCellWidthInChars(sheet, cell) - CELL_PADDING_CHAR_WIDTH);
		double charsPerLine = Math.floor(columnWidth * TEMPLATE_CHAR_WIDTH_FACTOR);
		return Math.max(4.0D, Math.min(LONG_TEXT_CHARS_PER_LINE, charsPerLine));
	}

	/**
	 * 使用Excel列宽字符数做保守估算。
	 * 像素字体回退在服务器环境可能偏差较大，因此这里再按中文字符宽度兜底，取两种算法的最大值。
	 * @param sheet
	 * @param cell
	 * @param lines
	 * @return
	 */
	private int estimateCellLineCountByChars(Sheet sheet, Cell cell, String[] lines) {
		double columnWidth = Math.max(1.0D, (this.getCellWidthInChars(sheet, cell) - CELL_PADDING_CHAR_WIDTH) * TEMPLATE_CHAR_WIDTH_FACTOR);
		int lineCount = 0;
		for (String line : lines) {
			lineCount += this.estimateWrappedLineCountByChars(line, columnWidth);
		}

		return lineCount;
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
	 * 得到单元格可用列宽像素值，合并单元格按合并区域总宽度估算。
	 * @param sheet
	 * @param cell
	 * @return
	 */
	private double getCellWidthInPixels(Sheet sheet, Cell cell) {
		CellRangeAddress region = this.getMergedRegion(sheet, cell.getRowIndex(), cell.getColumnIndex());
		if (region==null) {
			return sheet.getColumnWidthInPixels(cell.getColumnIndex());
		}

		double width = 0.0D;
		for (int i=region.getFirstColumn(); i<=region.getLastColumn(); i++) {
			width += sheet.getColumnWidthInPixels(i);
		}

		return width;
	}

	/**
	 * 得到单元格可用列宽字符数，合并单元格按合并区域总宽度估算。
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
	 * 按列宽字符数保守估算一段无显式换行文本会被折成几行。
	 * @param text
	 * @param columnWidth
	 * @return
	 */
	private int estimateWrappedLineCountByChars(String text, double columnWidth) {
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
	 * 粗略估算字符宽度，中文/日文/韩文及中文标点按2个英文字符计算。
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
		Character.UnicodeBlock block = Character.UnicodeBlock.of(ch);
		if (block==Character.UnicodeBlock.CJK_SYMBOLS_AND_PUNCTUATION
				|| block==Character.UnicodeBlock.HALFWIDTH_AND_FULLWIDTH_FORMS
				|| block==Character.UnicodeBlock.GENERAL_PUNCTUATION) {
			return 2.0D;
		}
		if (ch>255) {
			return 1.5D;
		}

		return 1.0D;
	}

	/**
	 * 从单元格样式中提取字体信息，用于按真实像素宽度估算换行。
	 * AWT没有对应字体时会自动回退到可用字体，仍比按字符个数估算更接近Excel渲染。
	 * @param sheet
	 * @param cell
	 * @return
	 */
	private java.awt.Font getAwtFont(Sheet sheet, Cell cell) {
		String fontName = "Dialog";
		int fontSize = 11;
		int fontStyle = java.awt.Font.PLAIN;

		try {
			CellStyle cellStyle = cell.getCellStyle();
			org.apache.poi.ss.usermodel.Font poiFont = sheet.getWorkbook().getFontAt(cellStyle.getFontIndexAsInt());
			return this.toAwtFont(poiFont);
		} catch (Exception e) {
			// 字体信息读取失败时使用默认字体继续估算，避免模板填充被非关键样式问题中断。
		}

		return new java.awt.Font(fontName, fontStyle, Math.max(1, fontSize));
	}

	/**
	 * 获取工作簿默认字体，作为AWT行高到Excel行高的校正基准。
	 * @param sheet
	 * @return
	 */
	private java.awt.Font getWorkbookDefaultAwtFont(Sheet sheet) {
		try {
			return this.toAwtFont(sheet.getWorkbook().getFontAt(0));
		} catch (Exception e) {
			// 默认字体读取失败时使用Dialog 11号继续估算，保证模板填充不中断。
			return new java.awt.Font("Dialog", java.awt.Font.PLAIN, 11);
		}
	}

	/**
	 * 将POI字体转换为AWT字体，用于真实字体宽高测量。
	 * @param poiFont
	 * @return
	 */
	private java.awt.Font toAwtFont(org.apache.poi.ss.usermodel.Font poiFont) {
		String fontName = "Dialog";
		int fontSize = 11;
		int fontStyle = java.awt.Font.PLAIN;
		if (poiFont!=null) {
			if (poiFont.getFontName()!=null && poiFont.getFontName().trim().length()>0) {
				fontName = poiFont.getFontName();
			}
			if (poiFont.getFontHeightInPoints()>0) {
				fontSize = poiFont.getFontHeightInPoints();
			}
			if (poiFont.getBold()) {
				fontStyle = fontStyle | java.awt.Font.BOLD;
			}
			if (poiFont.getItalic()) {
				fontStyle = fontStyle | java.awt.Font.ITALIC;
			}
		}

		return new java.awt.Font(fontName, fontStyle, Math.max(1, fontSize));
	}

	/**
	 * 设置模板行高，保留模板已有高度或图片撑开的高度，不会把行高调小。
	 * @param sheet
	 * @param firstRow
	 * @param rowSpan
	 * @param lineCount
	 */
	private void setAutoTemplateRowHeight(Sheet sheet, int firstRow, int rowSpan, float targetTotalHeight) {
		float baseHeight = sheet.getDefaultRowHeightInPoints();
		if (baseHeight<=0) {
			baseHeight = 15.0F;
		}
		float currentTotalHeight = 0.0F;
		for (int i=0; i<rowSpan; i++) {
			Row row = sheet.getRow(firstRow + i);
			currentTotalHeight += row==null ? baseHeight : row.getHeightInPoints();
		}
		if (currentTotalHeight>=targetTotalHeight) {
			return;
		}

		float targetRowHeight = targetTotalHeight / rowSpan;
		for (int i=0; i<rowSpan; i++) {
			Row row = sheet.getRow(firstRow + i);
			if (row==null) {
				row = sheet.createRow(firstRow + i);
			}
			targetRowHeight = Math.min(targetRowHeight, MAX_EXCEL_ROW_HEIGHT_POINTS);
			if (targetRowHeight>row.getHeightInPoints()) {
				row.setHeightInPoints(targetRowHeight);
			}
		}
	}

	/**
	 * 记录列表模板单元格信息。
	 * 需要保留完整模板文本，避免共享单元格只替换第一个列表占位符。
	 */
	/**
	 * 计算同一个list占位符对应的模板逻辑块。
	 * JES通常是一条数据一行；TIS可能一条数据跨多行或包含纵向合并单元格，不能用单个占位符自己的rowSpan推导下一条数据位置。
	 * @param sheet
	 * @param templateListCells
	 * @return
	 */
	private TemplateListLayout getTemplateListLayout(Sheet sheet, List<TemplateListCell> templateListCells) {
		int startRow = Integer.MAX_VALUE;
		int endRow = -1;
		for (TemplateListCell templateListCell : templateListCells) {
			startRow = Math.min(startRow, templateListCell.getRowIndex());
			int rowSpan = this.getTemplateCellRowSpan(sheet, templateListCell);
			endRow = Math.max(endRow, templateListCell.getRowIndex() + rowSpan - 1);
		}
		if (startRow==Integer.MAX_VALUE || endRow<startRow) {
			return new TemplateListLayout(0, 1);
		}

		return new TemplateListLayout(startRow, endRow - startRow + 1);
	}

	/**
	 * 获取模板占位符自身覆盖的行数。
	 * 只有位于合并区域左上角的占位符才按合并行数计算，避免误把合并区域内部的普通空单元格当成新的数据起点。
	 * @param sheet
	 * @param templateListCell
	 * @return
	 */
	private int getTemplateCellRowSpan(Sheet sheet, TemplateListCell templateListCell) {
		CellRangeAddress mergedRegion = this.getMergedRegion(sheet, templateListCell.getRowIndex(), templateListCell.getColIndex());
		if (mergedRegion!=null
				&& mergedRegion.getFirstRow()==templateListCell.getRowIndex()
				&& mergedRegion.getFirstColumn()==templateListCell.getColIndex()) {
			return Math.max(1, mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1);
		}

		return 1;
	}

	/**
	 * list模板的一条数据在Sheet中占用的逻辑行块。
	 */
	private static class TemplateListLayout {
		private final int startRow;
		private final int blockRows;

		private TemplateListLayout(int startRow, int blockRows) {
			this.startRow = startRow;
			this.blockRows = Math.max(1, blockRows);
		}

		private int getStartRow() {
			return startRow;
		}

		private int getBlockRows() {
			return blockRows;
		}
	}

	/**
	 * 记录列表模板单元格信息，保留完整模板文本以支持同一单元格内多个占位符替换。
	 */
	private static class TemplateListCell {
		private final String listKey;
		private final int rowIndex;
		private final int colIndex;
		private final String cellTemplate;
		private final List<String> placeholders;
		private final boolean exclusive;
		private final String exclusiveAttributeKey;

		private TemplateListCell(String listKey, int rowIndex, int colIndex, String cellTemplate, List<String> placeholders) {
			this.listKey = listKey;
			this.rowIndex = rowIndex;
			this.colIndex = colIndex;
			this.cellTemplate = cellTemplate;
			this.placeholders = new ArrayList<String>(placeholders);

			String exclusivePlaceholder = listKey + ".";
			this.exclusive = placeholders.size()==1
					&& cellTemplate.equals("${" + placeholders.get(0) + "}")
					&& placeholders.get(0).startsWith(exclusivePlaceholder);
			this.exclusiveAttributeKey = this.exclusive ? placeholders.get(0).substring(exclusivePlaceholder.length()) : null;
		}

		private String getListKey() {
			return listKey;
		}

		private int getRowIndex() {
			return rowIndex;
		}

		private int getColIndex() {
			return colIndex;
		}

		private String getCellTemplate() {
			return cellTemplate;
		}

		private List<String> getPlaceholders() {
			return placeholders;
		}

		private boolean isExclusive() {
			return exclusive;
		}

		private String getExclusiveAttributeKey() {
			return exclusiveAttributeKey;
		}
	}

	/**
	 * 转成Map类型
	 * @param obj
	 * @return
	 */
	@SuppressWarnings("unchecked")
	private Map<String, Object> convertToMap(Object obj) {
		if (obj instanceof Map) {
			return (Map<String, Object>) obj;
		}

		Map<String, Object> map = new HashMap<>();
		try {
			for (PropertyDescriptor pd : Introspector.getBeanInfo(obj.getClass(), Object.class).getPropertyDescriptors()) {
				Method getter = pd.getReadMethod();
				if (getter != null) {
					Object value = getter.invoke(obj);
					map.put(pd.getName(), value);
				}
			}
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		return map;
	}

	/**
	 * 复制模板
	 * @param sheet
	 * @param templateRows
	 * @param copyTimes
	 * @param makePageBreakByBlock
	 */
	@Override
	public void copyTemplate(Sheet sheet, int templateRows, int copyTimes, boolean makePageBreakByBlock) {
		RowHelper rowHelper = new RowHelper();
		CellRangeAddress originalPrintArea = getCurrentPrintArea(sheet);
		int srcStartRow = originalPrintArea==null ? 0 : originalPrintArea.getFirstRow();
		int blockRows = getTemplateBlockRows(sheet, srcStartRow, templateRows, originalPrintArea);
		if (blockRows<=0) {
			return;
		}

		// 适用于所有复制的偏移量
		for (int n = 1; n <= copyTimes; n++) {
			int destStartRow = srcStartRow + n * blockRows;
			// 1. 复制每一行
			for (int i = 0; i < blockRows; i++) {
				Row srcRow = sheet.getRow(srcStartRow + i);
				Row tgtRow = sheet.createRow(destStartRow + i);
				rowHelper.copyRow(sheet.getWorkbook(), srcRow, tgtRow);
			}
			// 2. 合并单元格
			for (int i = 0, num = sheet.getNumMergedRegions(); i < num; i++) {
				CellRangeAddress cra = sheet.getMergedRegion(i);
				// 只复制模板区的
				if (cra.getFirstRow() >= srcStartRow && cra.getLastRow() < srcStartRow + blockRows) {
					CellRangeAddress newCra = new CellRangeAddress(
							cra.getFirstRow() - srcStartRow + destStartRow,
							cra.getLastRow() - srcStartRow + destStartRow,
							cra.getFirstColumn(),
							cra.getLastColumn());
					sheet.addMergedRegion(newCra);
				}
			}
			// 3. 复制图片
			copyPictures(sheet, sheet, srcStartRow, srcStartRow + blockRows, destStartRow - srcStartRow);
		}

		// 4. 统一设置分页和打印区域
		if (makePageBreakByBlock) {
			int totalRows = (copyTimes + 1) * blockRows; // 模板+N次复制，每个都是一个打印区域块
			makePageBreakByBlock(sheet, srcStartRow, blockRows, totalRows, originalPrintArea, copyTimes + 1);
		}
	}

	/**
	 * 设置分页
	 * @param sheet
	 * @param firstBlockRow
	 * @param blockSize
	 * @param totalRows
	 */
	private static void makePageBreakByBlock(Sheet sheet, int firstBlockRow, int blockSize, int totalRows, CellRangeAddress originalPrintArea, int pageCount) {
		// 1. 清除已存在的全部分页符
		int[] breaks = sheet.getRowBreaks();
		for (int br : breaks) {
			sheet.removeRowBreak(br);
		}
		// 2. 每页blockSize，循环设置分页符
		for (int i = firstBlockRow + blockSize - 1; i < firstBlockRow + totalRows - 1; i += blockSize) {
			sheet.setRowBreak(i);
		}
		// 3. 打印区域
		int firstCol = 0;
		int lastCol = -1;
		int firstRow = firstBlockRow;
		int lastRow = firstBlockRow + totalRows - 1;
		if (originalPrintArea!=null) {
			// 模板本身已经设置了打印区域时，复制后必须沿用原来的左右边界。
			// 否则 getLastCellNum() 会把右侧曾经编辑过或带样式的空白列也纳入打印区域。
			firstCol = originalPrintArea.getFirstColumn();
			lastCol = originalPrintArea.getLastColumn();
			firstRow = originalPrintArea.getFirstRow();
			lastRow = firstRow + pageCount * blockSize - 1;
		} else {
			Row row = sheet.getRow(0);
			if (row==null || row.getLastCellNum()<1) {
				return;
			}
			lastCol = row.getLastCellNum() - 1;
		}
		Workbook workbook = sheet.getWorkbook();
		workbook.setPrintArea(workbook.getSheetIndex(sheet), firstCol, lastCol, firstRow, lastRow);
		keepWidthFitAndAllowVerticalPages(sheet);
	}

	/**
	 * 获取单个模板页块的行数。
	 * 优先使用模板已有的第一个横向分页符，因为它最能代表蓝色分页预览里的“当前页”边界；
	 * 没有分页符时再使用打印区域高度，最后才回退到调用方传入的templateRows。
	 * @param sheet
	 * @param srcStartRow
	 * @param templateRows
	 * @param originalPrintArea
	 * @return
	 */
	/**
	 * 复制模板后需要保留横向适配，但释放纵向页数限制。
	 * TIS模板在KKFile/LibreOffice预览时依赖fitWidth避免右侧截断；如果继续保留模板原来的fitHeight=1，
	 * 多个模板块会被强行压成一页，手动分页符就不会按预期生效。
	 * @param sheet
	 */
	private static void keepWidthFitAndAllowVerticalPages(Sheet sheet) {
		try {
			sheet.getPrintSetup().setFitHeight((short) 0);
		} catch (Exception e) {
			// 个别旧模板打印设置不完整时忽略该兼容性修正，避免复制模板中断。
		}
	}

	private static int getTemplateBlockRows(Sheet sheet, int srcStartRow, int templateRows, CellRangeAddress originalPrintArea) {
		int[] rowBreaks = sheet.getRowBreaks();
		Arrays.sort(rowBreaks);
		for (int rowBreak : rowBreaks) {
			if (rowBreak>=srcStartRow) {
				int blockRows = rowBreak - srcStartRow + 1;
				if (blockRows>0) {
					return blockRows;
				}
			}
		}

		if (originalPrintArea!=null) {
			return originalPrintArea.getLastRow() - originalPrintArea.getFirstRow() + 1;
		}

		return templateRows;
	}

	/**
	 * 获取当前Sheet已设置的打印区域。
	 * 模板复制前先保存原始区域，复制后才能准确保留模板设计的左右边界。
	 * @param sheet
	 * @return
	 */
	private static CellRangeAddress getCurrentPrintArea(Sheet sheet) {
		Workbook workbook = sheet.getWorkbook();
		String printArea = workbook.getPrintArea(workbook.getSheetIndex(sheet));
		if (printArea==null || printArea.trim().length()==0) {
			return null;
		}

		String areaRef = printArea.trim();
		int sheetNameEndIndex = areaRef.lastIndexOf('!');
		if (sheetNameEndIndex>=0) {
			areaRef = areaRef.substring(sheetNameEndIndex + 1);
		}
		int multiAreaIndex = areaRef.indexOf(',');
		if (multiAreaIndex>=0) {
			areaRef = areaRef.substring(0, multiAreaIndex);
		}

		try {
			return CellRangeAddress.valueOf(areaRef.replace("$", ""));
		} catch (Exception e) {
			// 打印区域格式异常时回退到旧逻辑，避免复制模板直接失败。
			return null;
		}
	}

	/**
	 * 根据手动分页块的实际宽高调整打印缩放。
	 * 自动行高会改变每页块的实际高度；这里直接计算需要的scale，确保每个手动分页块能放进一页，避免Excel在块内部再切自动分页。
	 * @param sheet
	 */
	private static void adjustPrintScaleForManualPageBreaks(Sheet sheet) {
		int[] rowBreaks = sheet.getRowBreaks();
		if (rowBreaks==null || rowBreaks.length==0) {
			return;
		}
		CellRangeAddress printArea = getCurrentPrintArea(sheet);
		if (printArea==null) {
			return;
		}

		double printableHeight = getPrintablePageHeightInPoints(sheet);
		double printableWidth = getPrintablePageWidthInPoints(sheet);
		if (printableHeight<=0 || printableWidth<=0) {
			return;
		}

		double maxBlockHeight = getMaxManualPageBlockHeightInPoints(sheet, printArea, rowBreaks);
		double printAreaWidth = getPrintAreaWidthInPoints(sheet, printArea);
		if (maxBlockHeight<=0 || printAreaWidth<=0) {
			return;
		}

		int heightScale = (int) Math.floor(printableHeight * 100.0D / maxBlockHeight);
		int widthScale = (int) Math.floor(printableWidth * 100.0D / printAreaWidth);
		int targetScale = Math.min(heightScale, widthScale);
		targetScale = Math.max(10, Math.min(400, targetScale));

		try {
			PrintSetup printSetup = sheet.getPrintSetup();
			short currentScale = printSetup.getScale();
			if (currentScale<=0) {
				currentScale = 100;
			}
			if (targetScale>=currentScale) {
				// 短文本或未明显撑高时不需要切换到固定scale；保留fitWidth，KKFile才能继续按一页宽预览。
				// 这里只释放纵向页数限制，避免复制模板后fitHeight=1把多页压成一页。
				keepWidthFitAndAllowVerticalPages(sheet);
				return;
			}

			int finalScale = Math.min((int) currentScale, targetScale);
			finalScale = Math.max(10, Math.min(400, finalScale));

			// 手动分页已经确定了页块边界，此处使用明确scale而不是fitHeight，避免Excel按旧缩放或自动分页重新拆页。
			sheet.setFitToPage(false);
			sheet.setAutobreaks(false);
			printSetup.setFitWidth((short) 0);
			printSetup.setFitHeight((short) 0);
			printSetup.setScale((short) finalScale);
			disableAutoPageBreaks(sheet);
			clearFitToPageForScale(sheet);
		} catch (Exception e) {
			// 个别旧模板可能没有完整打印设置，忽略scale修正，不影响打印区域扩展。
		}
	}

	/**
	 * 计算手动分页块中的最大实际高度。
	 * @param sheet
	 * @param printArea
	 * @param rowBreaks
	 * @return
	 */
	private static double getMaxManualPageBlockHeightInPoints(Sheet sheet, CellRangeAddress printArea, int[] rowBreaks) {
		Arrays.sort(rowBreaks);
		double maxHeight = 0.0D;
		int blockStartRow = printArea.getFirstRow();
		for (int rowBreak : rowBreaks) {
			if (rowBreak<printArea.getFirstRow()) {
				continue;
			}
			if (rowBreak>printArea.getLastRow()) {
				break;
			}
			maxHeight = Math.max(maxHeight, getRowsHeightInPoints(sheet, blockStartRow, rowBreak));
			blockStartRow = rowBreak + 1;
		}
		if (blockStartRow<=printArea.getLastRow()) {
			maxHeight = Math.max(maxHeight, getRowsHeightInPoints(sheet, blockStartRow, printArea.getLastRow()));
		}

		return maxHeight;
	}

	/**
	 * 计算指定行范围的实际总高度。
	 * @param sheet
	 * @param firstRow
	 * @param lastRow
	 * @return
	 */
	private static double getRowsHeightInPoints(Sheet sheet, int firstRow, int lastRow) {
		double height = 0.0D;
		float defaultHeight = sheet.getDefaultRowHeightInPoints();
		if (defaultHeight<=0) {
			defaultHeight = 15.0F;
		}
		for (int i=firstRow; i<=lastRow; i++) {
			Row row = sheet.getRow(i);
			height += row==null ? defaultHeight : row.getHeightInPoints();
		}

		return height;
	}

	/**
	 * 计算打印区域宽度，单位为磅。
	 * @param sheet
	 * @param printArea
	 * @return
	 */
	private static double getPrintAreaWidthInPoints(Sheet sheet, CellRangeAddress printArea) {
		double widthPixels = 0.0D;
		for (int i=printArea.getFirstColumn(); i<=printArea.getLastColumn(); i++) {
			widthPixels += sheet.getColumnWidthInPixels(i);
		}

		return widthPixels / POINTS_TO_PIXELS;
	}

	/**
	 * 获取当前纸张扣除上下页边距后的可打印高度，单位为磅。
	 * @param sheet
	 * @return
	 */
	private static double getPrintablePageHeightInPoints(Sheet sheet) {
		double[] pageSize = getPageSizeInPoints(sheet);
		double height = sheet.getPrintSetup().getLandscape() ? pageSize[0] : pageSize[1];
		return height - (sheet.getMargin(PageMargin.TOP) + sheet.getMargin(PageMargin.BOTTOM)) * 72.0D;
	}

	/**
	 * 获取当前纸张扣除左右页边距后的可打印宽度，单位为磅。
	 * @param sheet
	 * @return
	 */
	private static double getPrintablePageWidthInPoints(Sheet sheet) {
		double[] pageSize = getPageSizeInPoints(sheet);
		double width = sheet.getPrintSetup().getLandscape() ? pageSize[1] : pageSize[0];
		return width - (sheet.getMargin(PageMargin.LEFT) + sheet.getMargin(PageMargin.RIGHT)) * 72.0D;
	}

	/**
	 * 获取常见纸张尺寸，返回宽高，单位为磅。
	 * 未识别纸张时按A4处理，兼容多数模板。
	 * @param sheet
	 * @return
	 */
	private static double[] getPageSizeInPoints(Sheet sheet) {
		short paperSize = sheet.getPrintSetup().getPaperSize();
		if (paperSize==PrintSetup.LETTER_PAPERSIZE) {
			return new double[] {612.0D, 792.0D};
		}
		if (paperSize==PrintSetup.LEGAL_PAPERSIZE) {
			return new double[] {612.0D, 1008.0D};
		}
		if (paperSize==PrintSetup.A3_PAPERSIZE) {
			return new double[] {841.89D, 1190.55D};
		}
		if (paperSize==PrintSetup.A5_PAPERSIZE) {
			return new double[] {419.53D, 595.28D};
		}

		return new double[] {595.28D, 841.89D};
	}

	/**
	 * 关闭Excel自动分页。
	 * 模板复制后已经在模板块边界设置了手动分页符；如果自动分页仍开启，长文本撑高后Excel会在块内部再切出虚线分页。
	 * @param sheet
	 */
	private static void disableAutoPageBreaks(Sheet sheet) {
		if (!(sheet instanceof XSSFSheet)) {
			return;
		}
		try {
			XSSFSheet xssfSheet = (XSSFSheet) sheet;
			org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet ctWorksheet = xssfSheet.getCTWorksheet();
			if (!ctWorksheet.isSetSheetPr()) {
				ctWorksheet.addNewSheetPr();
			}
			if (!ctWorksheet.getSheetPr().isSetPageSetUpPr()) {
				ctWorksheet.getSheetPr().addNewPageSetUpPr();
			}
			ctWorksheet.getSheetPr().getPageSetUpPr().setFitToPage(false);
			ctWorksheet.getSheetPr().getPageSetUpPr().setAutoPageBreaks(false);
		} catch (Exception e) {
			// 底层XML写入失败时保留POI的setAutobreaks(false)，避免模板复制中断。
		}
	}

	/**
	 * 使用明确scale时，需要关闭底层fitToPage。
	 * @param sheet
	 */
	private static void clearFitToPageForScale(Sheet sheet) {
		if (!(sheet instanceof XSSFSheet)) {
			return;
		}
		try {
			XSSFSheet xssfSheet = (XSSFSheet) sheet;
			org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet ctWorksheet = xssfSheet.getCTWorksheet();
			if (!ctWorksheet.isSetSheetPr()) {
				ctWorksheet.addNewSheetPr();
			}
			if (!ctWorksheet.getSheetPr().isSetPageSetUpPr()) {
				ctWorksheet.getSheetPr().addNewPageSetUpPr();
			}
			ctWorksheet.getSheetPr().getPageSetUpPr().setFitToPage(false);
		} catch (Exception e) {
			// 底层fitToPage关闭失败时保留POI已设置的scale，避免模板复制中断。
		}
	}

	/**
	 * 复制图片，复制templateRows范围的图片到偏移的新行
	 * @param src
	 * @param tgt
	 * @param limitStart
	 * @param limitEnd
	 * @param rowOffset
	 */
	private static void copyPictures(Sheet src, Sheet tgt, int limitStart, int limitEnd, int rowOffset) {
		Workbook workbook = src.getWorkbook();
		if (!(src.getDrawingPatriarch() instanceof XSSFDrawing)) {
			return;
		}
		XSSFDrawing srcDraw = (XSSFDrawing) src.getDrawingPatriarch();
		Object tgtDrawing = tgt.createDrawingPatriarch();
		if (!(tgtDrawing instanceof XSSFDrawing)) {
			return;
		}
		XSSFDrawing tgtDraw = (XSSFDrawing) tgtDrawing;

		List<XSSFShape> shapes = srcDraw.getShapes();
		for (XSSFShape shape : shapes) {
			if (shape instanceof XSSFPicture) {
				XSSFPicture srcPic = (XSSFPicture) shape;
				XSSFClientAnchor anchor = (XSSFClientAnchor) srcPic.getAnchor();

				// 判断图片是否在复制行区间
				if (anchor.getRow1() >= limitStart && anchor.getRow2() < limitEnd) {
					int newRow1 = anchor.getRow1() + rowOffset;
					int newRow2 = anchor.getRow2() + rowOffset;
					XSSFClientAnchor newAnchor = new XSSFClientAnchor(
							anchor.getDx1(), anchor.getDy1(),
							anchor.getDx2(), anchor.getDy2(),
							anchor.getCol1(), newRow1,
							anchor.getCol2(), newRow2);
					int pictureIdx = workbook.addPicture(
							srcPic.getPictureData().getData(),
							srcPic.getPictureData().getPictureType()
					);
					tgtDraw.createPicture(newAnchor, pictureIdx);
				}
			}
		}
	}

	/**
	 * 复制Sheet
	 * @param workbook
	 * @param sourceSheetName    源Sheet名称
	 * @param targetSheetName    目标Sheet名称
	 * @param copyCount          复制次数
	 */
	@Override
	public void copySheets(Workbook workbook, String sourceSheetName, String targetSheetName, int copyCount, SheetPrintArea printArea) {
		// 1.校验
		if (workbook == null) {
			throw new IllegalArgumentException("Workbook is null");
		}
		if (copyCount <= 0) {
			return;
		}

		Sheet sourceSheet = workbook.getSheet(sourceSheetName);
		if (sourceSheet == null) {
			throw new RuntimeException("Source sheet '" + sourceSheetName + "' not found.");
		}

		// 2.复制sheet页
		for (int i = 0; i < copyCount; i++) {
			String newSheetName = targetSheetName + (i + 2);

			// 2.1.检查重名，删除旧Sheet
			int oldIndex = workbook.getSheetIndex(newSheetName);
			if (oldIndex != -1) {
				workbook.removeSheetAt(oldIndex);
			}

			// 2.2.创建新Sheet (使用 cloneSheet)
			Sheet newSheet = workbook.cloneSheet(workbook.getSheetIndex(sourceSheetName));
			int newSheetIndex = workbook.getSheetIndex(newSheet);
			workbook.setSheetName(newSheetIndex, newSheetName);

			// 2.3.打印区域
			if (printArea != null) {
				// 打印格式
				this.copyPrintSettings(sourceSheet, newSheet, printArea.getFirstRow(), printArea.getLastRow(), printArea.getFirstColumn(), printArea.getLastColumn());
				// 设置视图模式 (分页预览)
				this.copyViewMode(sourceSheet, newSheet);
			}

			// 2.4.设置缩放格式
			this.copyZoom(sourceSheet, newSheet);
		}
	}

	/**
	 * 根据指定起始/终止行（1-based）和列（字母，如"A"）设置打印区域并复制打印设置
	 *
	 * @param source      源Sheet
	 * @param target      目标Sheet
	 * @param firstRow    起始行，1-based，例如1表示Excel的第1行
	 * @param lastRow     终止行，1-based，包含
	 * @param firstColumn 起始列，如"A"
	 * @param lastColumn  终止列，如"F"
	 */
	private void copyPrintSettings(Sheet source, Sheet target,
			int firstRow, int lastRow, String firstColumn, String lastColumn) {
		if (!(source instanceof XSSFSheet) || !(target instanceof XSSFSheet)) {
			return;
		}
		XSSFSheet srcSheet = (XSSFSheet) source;
		XSSFSheet destSheet = (XSSFSheet) target;

		// 1.设置打印区域
		String areaRef = firstColumn.toUpperCase() + firstRow + ":" + lastColumn.toUpperCase() + lastRow;
		int destIndex = destSheet.getWorkbook().getSheetIndex(destSheet);
		destSheet.getWorkbook().setPrintArea(destIndex, areaRef);

		// 2.复制页眉页脚
		try {
			if (srcSheet.getHeader() != null) {
				destSheet.getHeader().setCenter(srcSheet.getHeader().getCenter());
				destSheet.getHeader().setLeft(srcSheet.getHeader().getLeft());
				destSheet.getHeader().setRight(srcSheet.getHeader().getRight());
			}
		} catch (Exception e) {
			// 某些模板没有完整页眉定义，忽略该项并继续复制其他打印设置。
		}
		try {
			if (srcSheet.getFooter() != null) {
				destSheet.getFooter().setCenter(srcSheet.getFooter().getCenter());
				destSheet.getFooter().setLeft(srcSheet.getFooter().getLeft());
				destSheet.getFooter().setRight(srcSheet.getFooter().getRight());
			}
		} catch (Exception e) {
			// 某些模板没有完整页脚定义，忽略该项并继续复制其他打印设置。
		}

		// 3.复制打印属性
		XSSFPrintSetup srcPrint = srcSheet.getPrintSetup();
		XSSFPrintSetup destPrint = destSheet.getPrintSetup();
		destPrint.setPaperSize(srcPrint.getPaperSize());
		destPrint.setLandscape(srcPrint.getLandscape());
		destPrint.setScale(srcPrint.getScale());
		destPrint.setFitWidth(srcPrint.getFitWidth());
		destPrint.setFitHeight(srcPrint.getFitHeight());
		destSheet.setPrintGridlines(srcSheet.isPrintGridlines());
		try { destPrint.setPageStart(srcPrint.getPageStart()); } catch (Exception e) {}
		try { destPrint.setDraft(srcPrint.getDraft()); } catch (Exception e) {}
		try { destPrint.setNoColor(srcPrint.getNoColor()); } catch (Exception e) {}
		try { destPrint.setNotes(srcPrint.getNotes()); } catch (Exception e) {}

		// 4.复制页边距
		for (PageMargin margin : PageMargin.values()) {
			destSheet.setMargin(margin, srcSheet.getMargin(margin));
		}

		// 5.复制重复标题行
		CellRangeAddress repeatingRows = srcSheet.getRepeatingRows();
		if (repeatingRows != null) {
			destSheet.setRepeatingRows(repeatingRows);
		}
	}

	/**
	 * 视图同步
	 * @param source
	 * @param target
	 */
	private void copyViewMode(Sheet source, Sheet target) {
		try {
			XSSFSheet srcSheet = (XSSFSheet) source;
			XSSFSheet destSheet = (XSSFSheet) target;
			// 1.【关键修复】清空目标文件所有旧的列分页符
			// 防止因为复制了源文件的旧设置，导致虚线位置不对
			// 我们只保留行分页符（如果源文件有分页），强制让列分页符由系统自动计算
			int[] existingColBreaks = destSheet.getColumnBreaks();
			for (int colBreak : existingColBreaks) {
				destSheet.removeColumnBreak(colBreak);
			}
			// 2.复制视图类型 (强制分页预览)
			CTSheetView srcView = srcSheet.getCTWorksheet().getSheetViews().getSheetViewArray(0);
			CTSheetView destView = destSheet.getCTWorksheet().getSheetViews().getSheetViewArray(0);

			if (srcView.isSetView()) {
				destView.setView(srcView.getView());
			}
			// 3.同步行分页符 (保留行分页，因为行分页通常是有意义的)
			int[] rowBreaks = srcSheet.getRowBreaks();
			for (int rowBreak : rowBreaks) {
				destSheet.setRowBreak(rowBreak);
			}
			// 4.强制同步页边距
			for (PageMargin margin : PageMargin.values()) {
				destSheet.setMargin(margin, srcSheet.getMargin(margin));
			}

			// 5.同步顶部显示的窗格位置
			if (srcView.isSetTopLeftCell()) {
				destView.setTopLeftCell(srcView.getTopLeftCell());
			}
		} catch (Exception e) {
			try {
				target.setAutobreaks(true);
			} catch (Exception ex) {
			}
		}
	}

	/**
	 * 复制 Sheet 的缩放比例 解决打开文件时视图大小不一致的问题
	 * @param source
	 * @param target
	 */
	private void copyZoom(Sheet source, Sheet target) {
		try {
			if (!(source instanceof XSSFSheet) || !(target instanceof XSSFSheet)) {
				return;
			}
			// 强制转换为 XSSF (因为 .xlsx 需要操作底层 XML)
			XSSFSheet srcSheet = (XSSFSheet) source;
			XSSFSheet destSheet = (XSSFSheet) target;

			// 获取源 Sheet 的视图对象
			CTSheetView srcView = srcSheet.getCTWorksheet().getSheetViews().getSheetViewArray(0);

			// 获取目标 Sheet 的视图对象
			CTSheetView destView = destSheet.getCTWorksheet().getSheetViews().getSheetViewArray(0);

			// 获取源 Sheet 的缩放比例
			// 注意：如果源文件是 100%，这里获取到的是 100
			long zoomScale = srcView.getZoomScale();
			// 如果源没有设置缩放（默认通常是100），getZoomScale 可能返回 null 或其他值
			// 这里做一个简单的保护，如果获取不到值，默认设为 100
			if (zoomScale <= 0) {
				zoomScale = 100;
			}
			// 设置目标 Sheet 的缩放比例
			destView.setZoomScale(zoomScale);
		} catch (Exception e) {
			// 如果发生异常（比如 XML 结构异常），不影响主流程。
		}
	}

}
