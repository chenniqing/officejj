package cn.javaex.officejj.excel.help;

import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.PageMargin;
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

import cn.javaex.officejj.excel.ExcelUtils;
import cn.javaex.officejj.excel.entity.SheetPrintArea;

/**
 * 模板替换写入Excel
 * 
 * @author 陈霓清
 */
public class SheetTemplateHelper extends SheetHelper {
	
	/**
	 * 替换占位符
	 */
	@Override
	public void write(Sheet sheet, Map<String, Object> param) {
		CellHelper cellHelper = new CellHelper();
		
		Map<String, List<Map<String, Integer[]>>> listMap = new LinkedHashMap<String, List<Map<String, Integer[]>>>();
		
		Row row = null;
		Cell cell = null;
		int index = 0;
		int lastRowNum = sheet.getLastRowNum();
		
		while (index <= lastRowNum) {
			row = sheet.getRow(index++);
			if (row==null) {
				continue;
			}
			
			List<Map<String, Integer[]>> list = new ArrayList<Map<String, Integer[]>>();
			String tempListKey = "";
			String listKey = "";
			
			int startCol = row.getFirstCellNum();    // 索引
			int endCol = row.getLastCellNum();       // 从1开始计算
			for (int i=startCol; i<endCol; i++) {
				if (row.getCell(i)==null) {
					continue;
				}
				
				// 得到单元格的内容
				String cellValue = ExcelUtils.getCellValue(row.getCell(i));
				
				// 如果单元格的内容不包含 ${xxx}，则跳过
				if (!(cellValue.contains("${") && cellValue.contains("}"))) {
					continue;
				}
				
				// 获取该单元格内的所有占位符变量
				List<String> placeholders = cellHelper.getPlaceholders(cellValue);
				
				// 如果是list遍历的话，一个格子中只能有一个占位符，且占位符中包含 “.” 符号
				// list遍历
				if (placeholders.get(0).contains(".")) {
					String[] arr = placeholders.get(0).split("\\.");
					listKey = arr[0];
					String attributeKey = arr[1];
					
					if (!"".equals(tempListKey) && !"".equals(listKey) && !tempListKey.equals(listKey)) {
						listMap.put(tempListKey, list);
						this.setListValue(sheet, listMap, param);
						
						tempListKey = "";
						list = new ArrayList<Map<String, Integer[]>>();
					} else {
						tempListKey = listKey;
					}
					
					Map<String, Integer[]> map = new HashMap<String, Integer[]>();
					map.put(attributeKey, new Integer[] {row.getRowNum(), i});
					
					list.add(map);
				}
				// 直接替换
				else {
					// 占位符独占一格时，需要根据替换值的实际类型进行替换
					if (cellValue.equals("${" + placeholders.get(0) + "}")) {
						cell = sheet.getRow(row.getRowNum()).getCell(i);
						cellHelper.setValue(cell, param.get(placeholders.get(0)));
						// 重置清空，当占位符一致时，只填充一个，以满足复制的场景
						param.put(placeholders.get(0), PLACEHOLDER_CLEAR);
					}
					// 占位符非独占一格时，认为该单元格的值是字符串，需要替换其中所有的占位符
					else {
						cell = sheet.getRow(row.getRowNum()).getCell(i);
						cellHelper.setValue(cell, placeholders, param);
					}
				}
			}
			
			if (!"".equals(listKey)) {
				listMap.put(listKey, list);
			}
		}
		
		if (listMap.isEmpty()==false) {
			this.setListValue(sheet, listMap, param);
		}
	}

	/**
	 * 替换模板中的占位符（list遍历）
	 * @param sheet
	 * @param listMap
	 * @param param
	 */
	@SuppressWarnings("unchecked")
	private void setListValue(Sheet sheet, Map<String, List<Map<String, Integer[]>>> listMap, Map<String, Object> param) {
		CellHelper cellHelper = new CellHelper();
		
		// LinkedHashMap倒序遍历
		ListIterator<Map.Entry<String, List<Map<String, Integer[]>>>> iterator = new ArrayList<Map.Entry<String, List<Map<String, Integer[]>>>>(listMap.entrySet()).listIterator(listMap.size());
		while (iterator.hasPrevious()) {
			Map.Entry<String, List<Map<String, Integer[]>>> entry = iterator.previous();
			
			// 1.0 取出需要遍历的list数据
			List<Map<String, Object>> list = (List<Map<String, Object>>) param.get(entry.getKey());
			if (list==null || list.isEmpty()) {
				continue;
			}
			
			// 2.0 遍历取出每一条数据并设置值
			int len = list.size();
			List<Map<String, Integer[]>> placeholders = entry.getValue();
			for (int i=0; i<len; i++) {
				Map<String, Object> dataMap = this.convertToMap(list.get(i));
				
				for (Map<String, Integer[]> placeholder : placeholders) {
					// 获取该行的每一个单元格记录
					for (Map.Entry<String, Integer[]> placeholderMap : placeholder.entrySet()) {
						String attributeKey = placeholderMap.getKey();         // 属性Key
						Integer[] coordinate = placeholderMap.getValue();      // Cell坐标（行索引、列索引）
						
						// 或直接获取合并行数，如果不是合并区左上角，返回1
						int rowSpan = super.getMergedRowSpanIfTopLeft(sheet, coordinate[0], coordinate[1]);
						
						// 设置值替换
						int rowIndex = coordinate[0] + i * rowSpan;
						int colIndex = coordinate[1];
						Row row = sheet.getRow(rowIndex);
						if (row == null) row = sheet.createRow(rowIndex);
						Cell cell = row.getCell(colIndex);
						if (cell == null) cell = row.createCell(colIndex);
						
						cellHelper.setValue(cell, dataMap.get(attributeKey));
					}
				}
			}
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
		
		// 适用于所有复制的偏移量
		for (int n = 1; n <= copyTimes; n++) {
			int srcStartRow = 0; // 模板区为0~26行（一般情况下）
			int destStartRow = n * templateRows; // 依次在[27, 54, 81, ...]行开头插入新块
			// 1. 复制每一行
			for (int i = 0; i < templateRows; i++) {
				Row srcRow = sheet.getRow(srcStartRow + i);
				Row tgtRow = sheet.createRow(destStartRow + i);
				rowHelper.copyRow(sheet.getWorkbook(), srcRow, tgtRow);
			}
			// 2. 合并单元格
			for (int i = 0, num = sheet.getNumMergedRegions(); i < num; i++) {
				CellRangeAddress cra = sheet.getMergedRegion(i);
				// 只复制模板区的
				if (cra.getFirstRow() >= srcStartRow && cra.getLastRow() < srcStartRow + templateRows) {
					CellRangeAddress newCra = new CellRangeAddress(
							cra.getFirstRow() - srcStartRow + destStartRow,
							cra.getLastRow() - srcStartRow + destStartRow,
							cra.getFirstColumn(),
							cra.getLastColumn());
					sheet.addMergedRegion(newCra);
				}
			}
			// 3. 复制图片
			copyPictures(sheet, sheet, srcStartRow, srcStartRow + templateRows, destStartRow - srcStartRow);
		}
		
		// 4. 统一设置分页和打印区域
		if (makePageBreakByBlock) {
			int totalRows = (copyTimes + 1) * templateRows; // 模板+N次复制，每个都是templateRows行
			makePageBreakByBlock(sheet, templateRows, totalRows);
		}
	}
	
	/**
	 * 设置分页
	 * @param sheet
	 * @param blockSize
	 * @param totalRows
	 */
	private static void makePageBreakByBlock(Sheet sheet, int blockSize, int totalRows) {
		// 1. 清除已存在的全部分页符
		int[] breaks = sheet.getRowBreaks();
		for (int br : breaks) {
			sheet.removeRowBreak(br);
		}
		// 2. 每页blockSize，循环设置分页符
		for (int i = blockSize-1; i < totalRows-1; i += blockSize) {
			sheet.setRowBreak(i);
		}
		// 3. 打印区域
		int firstCol = 0;
		int lastCol = sheet.getRow(0).getLastCellNum() - 1;
		Workbook workbook = sheet.getWorkbook();
		workbook.setPrintArea(workbook.getSheetIndex(sheet), firstCol, lastCol, 0, totalRows-1);
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
		XSSFDrawing srcDraw = (XSSFDrawing) src.getDrawingPatriarch();
		if (srcDraw == null) {
			return;
		}
		XSSFDrawing tgtDraw = (XSSFDrawing) tgt.createDrawingPatriarch();
		
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
		} catch (Exception e) { e.printStackTrace(); }
		try {
			if (srcSheet.getFooter() != null) {
				destSheet.getFooter().setCenter(srcSheet.getFooter().getCenter());
				destSheet.getFooter().setLeft(srcSheet.getFooter().getLeft());
				destSheet.getFooter().setRight(srcSheet.getFooter().getRight());
			}
		} catch (Exception e) { e.printStackTrace(); }
		
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
			// 如果发生异常（比如 XML 结构异常），不影响主流程，仅打印提示
			e.printStackTrace();
		}
	}
	
}
