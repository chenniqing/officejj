package cn.javaex.officejj.excel.help;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPresetLineDashProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTransform2D;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineEndLength;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineEndType;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineEndWidth;
import org.openxmlformats.schemas.drawingml.x2006.main.STPresetLineDashVal;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;

import cn.javaex.officejj.common.util.ArrayHandler;
import cn.javaex.officejj.excel.entity.ExcelSetting;
import cn.javaex.officejj.excel.entity.SheetPrintArea;

/**
 * Sheet操作
 * 
 * @author 陈霓清
 */
public class SheetHelper {
	
	/** 默认sheet页名称 */
	public static final String SHEET_NAME = "Sheet1";
	/** 行高基数 */
	public static final int BASE_ROW_HEIGHT = 20;
	/** 列宽基数 */
	public static final int BASE_COLUMN_WIDTH = 256;
	/** 重置清空，当占位符一致时，只填充一个，以满足复制的场景 */
	public static final String PLACEHOLDER_CLEAR = "==PLACEHOLDER_CLEAR==";
	
	// 存储值替换
	public Map<String, Object> replaceMap = new HashMap<String, Object>();
	// 存储格式化
	public Map<String, Object> formatMap = new HashMap<String, Object>();
	// 存储合并多个单元格数据的成员变量
	public Map<String, String> skipMap = new HashMap<String, String>();
	
	/**
	 * 自动合并列
	 * @param sheet
	 * @param colNum      第几列（从0开始计算）
	 * @param firstRow    起始行（从0开始计算）
	 * @param lastRow     终止行（从0开始计算）
	 * @param clazz
	 */
	public void setAutoMergeCol(Sheet sheet, int colNum, int firstRow, int lastRow, Class<?> clazz) {
		
	}
	
	/**
	 * 自动合并行
	 * @param sheet
	 * @param rowNum      第几行（从0开始计算）
	 * @param firstCol    起始列（从0开始计算）
	 * @param lastCol     终止列（从0开始计算）
	 * @param clazz
	 */
	public void setAutoMergeRow(Sheet sheet, int rowNum, int firstCol, Integer lastCol, Class<?> clazz) {
		
	}
	
	/**
	 * 设置合并
	 * @param sheet
	 * @param firstRow    起始行（从0开始计算）
	 * @param lastRow     终止行（从0开始计算）
	 * @param firstCol    起始列（从0开始计算）
	 * @param lastCol     终止列（从0开始计算）
	 * @param clazz
	 */
	public void setMerge(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, Class<?> clazz) {
		
	}
	
	/**
	 * 画线条，带收缩率
	 *
	 * @param colA 起点列，从0开始
	 * @param rowA 起点行，从0开始
	 * @param colB 终点列，从0开始
	 * @param rowB 终点行，从0开始
	 * @param colorRGB 颜色， int数组{r, g, b} 例如：new int[]{0, 0, 0}
	 * @param lineWidth 粗细，单位为磅
	 * @param shrinkRatio 收缩率，例如：填0.8，以防止遮挡单元格内容
	 * @param hasArrow 是否带箭头
	 * @param isDottedLine 是否是虚线，false为实线
	 */
	public void drawLine(Sheet sheet, int colA, int rowA, int colB, int rowB, 
			int[] colorRGB, double lineWidth, double shrinkRatio, boolean hasArrow, boolean isDottedLine) {
		// 1. anchor区范围
		int col1 = Math.min(colA, colB);
		int col2 = Math.max(colA, colB);
		int row1 = Math.min(rowA, rowB);
		int row2 = Math.max(rowA, rowB);
		
		// 2. 所有相关格子的cell左上EMU
		long[] colStarts = new long[Math.max(colA, colB)+2]; // 用于缓存提高效率
		colStarts[0] = 0;
		for(int i = 1; i < colStarts.length; i++) {
			colStarts[i] = colStarts[i-1] + Units.pixelToEMU((int) sheet.getColumnWidthInPixels(i-1));
		}
		
		long[] rowStarts = new long[Math.max(rowA, rowB) + 2];
		rowStarts[0] = 0;
		for (int i=1;i<rowStarts.length;i++) {
			Row rObj = sheet.getRow(i-1);
			float rh = rObj != null ? rObj.getHeightInPoints() : sheet.getDefaultRowHeightInPoints();
			rowStarts[i] = rowStarts[i-1] + Units.pixelToEMU(Math.round(rh * Units.PIXEL_DPI / Units.POINT_DPI));
		}
		
		// 3. 计算两点实际中心，缩短80%，得到新起点(sx, sy)和新终点(ex, ey)
		int cell1w = (int) sheet.getColumnWidthInPixels(colA);
		int cell1h = Math.round((sheet.getRow(rowA)!=null ? sheet.getRow(rowA).getHeightInPoints() : sheet.getDefaultRowHeightInPoints()) * Units.PIXEL_DPI / Units.POINT_DPI);
		int cell2w = (int) sheet.getColumnWidthInPixels(colB);
		int cell2h = Math.round((sheet.getRow(rowB)!=null ? sheet.getRow(rowB).getHeightInPoints() : sheet.getDefaultRowHeightInPoints()) * Units.PIXEL_DPI / Units.POINT_DPI);
		
		long absCx1 = colStarts[colA] + Units.pixelToEMU(cell1w/2);
		long absCy1 = rowStarts[rowA] + Units.pixelToEMU(cell1h/2);
		long absCx2 = colStarts[colB] + Units.pixelToEMU(cell2w/2);
		long absCy2 = rowStarts[rowB] + Units.pixelToEMU(cell2h/2);
		
		double leave = (1 - shrinkRatio) / 2;
		long sx = (long)(absCx1 + (absCx2 - absCx1) * leave);
		long sy = (long)(absCy1 + (absCy2 - absCy1) * leave);
		long ex = (long)(absCx1 + (absCx2 - absCx1) * (1 - leave));
		long ey = (long)(absCy1 + (absCy2 - absCy1) * (1 - leave));
		
		// 4. 计算anchor左上、右下单元格的左上绝对EMU
		long anchorBaseX1 = colStarts[col1];
		long anchorBaseY1 = rowStarts[row1];
		long anchorBaseX2 = colStarts[col2];
		long anchorBaseY2 = rowStarts[row2];
		
		// 5. 计算dx1/dy1/dx2/dy2：均为“端点EMU-对应anchor格左上EMU”
		int dx1 = (int)(sx - anchorBaseX1);
		int dy1 = (int)(sy - anchorBaseY1);
		int dx2 = (int)(ex - anchorBaseX2);
		int dy2 = (int)(ey - anchorBaseY2);
		
		// 6. 安全限定：dx1/dy1 必须在 [0, anchor单元格宽/高EMU]内
		int anchor1wEMU = Units.pixelToEMU((int)sheet.getColumnWidthInPixels(col1));
		int anchor1hEMU = Units.pixelToEMU(Math.round((sheet.getRow(row1)!=null ? sheet.getRow(row1).getHeightInPoints() : sheet.getDefaultRowHeightInPoints()) * Units.PIXEL_DPI / Units.POINT_DPI));
		int anchor2wEMU = Units.pixelToEMU((int)sheet.getColumnWidthInPixels(col2));
		int anchor2hEMU = Units.pixelToEMU(Math.round((sheet.getRow(row2)!=null ? sheet.getRow(row2).getHeightInPoints() : sheet.getDefaultRowHeightInPoints()) * Units.PIXEL_DPI / Units.POINT_DPI));
		
		dx1 = Math.max(0, Math.min(dx1, anchor1wEMU));
		dy1 = Math.max(0, Math.min(dy1, anchor1hEMU));
		dx2 = Math.max(0, Math.min(dx2, anchor2wEMU));
		dy2 = Math.max(0, Math.min(dy2, anchor2hEMU));
		
		// 7. 创建anchor
		XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
		
		XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
		XSSFSimpleShape line = drawing.createSimpleShape(anchor);
		line.setShapeType(ShapeTypes.LINE);
		line.setLineWidth(lineWidth);
		line.setLineStyleColor(colorRGB[0], colorRGB[1], colorRGB[2]);
		
		// 线条
		CTShape ctShape = (CTShape) line.getCTShape();
		CTShapeProperties props = ctShape.getSpPr();
		if (props.isSetLn()) {
			CTLineProperties ln = props.getLn();
			// 设置虚线线型
			if (isDottedLine) {
				CTPresetLineDashProperties dash = ln.isSetPrstDash() ? ln.getPrstDash() : ln.addNewPrstDash();
				dash.setVal(STPresetLineDashVal.DASH);
			}
		}
		// 箭头
		if (hasArrow) {
			CTShapeProperties props2 = ctShape.getSpPr();
			if (props2.isSetLn()) {
				CTLineProperties ln = props2.getLn();
				if (colA > colB) {
					ln.addNewHeadEnd().setType(STLineEndType.TRIANGLE);
					ln.getHeadEnd().setW(STLineEndWidth.MED);
					ln.getHeadEnd().setLen(STLineEndLength.MED);
				} else {
					ln.addNewTailEnd().setType(STLineEndType.TRIANGLE);
					ln.getTailEnd().setW(STLineEndWidth.MED);
					ln.getTailEnd().setLen(STLineEndLength.MED);
				}
			}
		}
		
		// 垂直翻转
		if (colA > colB) {
			CTTransform2D xfrm = ctShape.getSpPr().getXfrm();
			xfrm.setFlipV(true);
		}
	}
	
	/**
	 * 画线条
	 *
	 * @param colA 起点列，从0开始
	 * @param rowA 起点行，从0开始
	 * @param colB 终点列，从0开始
	 * @param rowB 终点行，从0开始
	 * @param colorRGB 颜色， int数组{r, g, b} 例如：new int[]{0, 0, 0}
	 * @param lineWidth 粗细，单位为磅
	 * @param hasArrow 是否带箭头
	 * @param isDottedLine 是否是虚线，false为实线
	 */
	public void drawLine(Sheet sheet, int colA, int rowA, int colB, int rowB, int[] colorRGB, 
			double lineWidth, boolean hasArrow, boolean isDottedLine) {
		// anchor参数需满足 col1<=col2, row1<=row2
		int col1 = Math.min(colA, colB);
		int col2 = Math.max(colA, colB);
		int row1 = Math.min(rowA, rowB);
		int row2 = Math.max(rowA, rowB);
		
		// 获取实际宽高
		int cell1w = (int) sheet.getColumnWidthInPixels(colA);
		int cell1h = Math.round(sheet.getRow(rowA).getHeightInPoints() * Units.PIXEL_DPI / Units.POINT_DPI);
		int cell2w = (int) sheet.getColumnWidthInPixels(colB);
		int cell2h = Math.round(sheet.getRow(rowB).getHeightInPoints() * Units.PIXEL_DPI / Units.POINT_DPI);
		if (cell1h == 0) {
			cell1h = (int) (sheet.getDefaultRowHeightInPoints() * Units.PIXEL_DPI / Units.POINT_DPI);
		}
		if (cell2h == 0) {
			cell2h = (int) (sheet.getDefaultRowHeightInPoints() * Units.PIXEL_DPI / Units.POINT_DPI);
		}
		
		// 计算相对偏移，把线端分别对准目标格子中心
		int dx1, dy1, dx2, dy2;
		if (colA <= colB && rowA <= rowB) {
			// (colA,rowA) (左上) 到 (colB,rowB) (右下)
			dx1 = Units.pixelToEMU(cell1w/2);
			dy1 = Units.pixelToEMU(cell1h/2);
			dx2 = Units.pixelToEMU(cell2w/2);
			dy2 = Units.pixelToEMU(cell2h/2);
		} else if (colA > colB && rowA <= rowB) {
			// (colB,rowA) (左上) 到 (colA,rowB) (右下)
			dx1 = Units.pixelToEMU(cell2w/2);
			dy1 = Units.pixelToEMU(cell1h/2);
			dx2 = Units.pixelToEMU(cell1w/2);
			dy2 = Units.pixelToEMU(cell2h/2);
		} else if (colA <= colB && rowA > rowB) {
			dx1 = Units.pixelToEMU(cell1w/2);
			dy1 = Units.pixelToEMU(cell2h/2);
			dx2 = Units.pixelToEMU(cell2w/2);
			dy2 = Units.pixelToEMU(cell1h/2);
		} else {
			// (colB,rowB) (左上) 到 (colA,rowA) (右下)
			dx1 = Units.pixelToEMU(cell2w/2);
			dy1 = Units.pixelToEMU(cell2h/2);
			dx2 = Units.pixelToEMU(cell1w/2);
			dy2 = Units.pixelToEMU(cell1h/2);
		}
		
		// anchor范围必须col1<=col2，row1<=row2
		XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
		
		XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
		XSSFSimpleShape line = drawing.createSimpleShape(anchor);
		line.setShapeType(ShapeTypes.LINE);
		line.setLineWidth(lineWidth);
		line.setLineStyleColor(colorRGB[0], colorRGB[1], colorRGB[2]);
		
		// 线条
		CTShape ctShape = (CTShape) line.getCTShape();
		CTShapeProperties props = ctShape.getSpPr();
		if (props.isSetLn()) {
			CTLineProperties ln = props.getLn();
			// 设置虚线线型
			if (isDottedLine) {
				CTPresetLineDashProperties dash = ln.isSetPrstDash() ? ln.getPrstDash() : ln.addNewPrstDash();
				dash.setVal(STPresetLineDashVal.DASH);
			}
		}
		// 箭头
		if (hasArrow) {
			CTShapeProperties props2 = ctShape.getSpPr();
			if (props2.isSetLn()) {
				CTLineProperties ln = props2.getLn();
				if (colA > colB) {
					ln.addNewHeadEnd().setType(STLineEndType.TRIANGLE);
					ln.getHeadEnd().setW(STLineEndWidth.MED);
					ln.getHeadEnd().setLen(STLineEndLength.MED);
				} else {
					ln.addNewTailEnd().setType(STLineEndType.TRIANGLE);
					ln.getTailEnd().setW(STLineEndWidth.MED);
					ln.getTailEnd().setLen(STLineEndLength.MED);
				}
			}
		}
		
		// 垂直翻转
		if (colA > colB) {
			CTTransform2D xfrm = ctShape.getSpPr().getXfrm();
			xfrm.setFlipV(true);
		}
	}

	/**
	 * 复制模板
	 * @param sheet
	 * @param templateRows
	 * @param copyTimes
	 * @param makePageBreakByBlock 
	 */
	public void copyTemplate(Sheet sheet, int templateRows, int copyTimes, boolean makePageBreakByBlock) {
		
	}
	
	/**
	 * 创建Header
	 * @param sheet
	 * @param clazz
	 * @param title
	 * @throws Exception
	 */
	public void setHeader(Sheet sheet, Class<?> clazz, String title) throws Exception {
		
	}
	
	/**
	 * 设置基本属性
	 * @param sheet
	 * @param clazz
	 */
	public void setBasicData(Sheet sheet, Class<?> clazz) {
		
	}
	
	/**
	 * 根据注解创建内容
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @param title
	 * @throws Exception 
	 */
	public void write(Sheet sheet, Class<?> clazz, List<?> list, String title) throws Exception {
		
	}

	/**
	 * 根据注解创建内容
	 * 多线程
	 * @param sheet
	 * @param clazz
	 * @param list
	 * @throws Exception
	 */
	public void writeByThreads(Sheet sheet, Class<?> clazz, List<?> list) throws Exception {
		
	}
	
	/**
	 * 根据设置类创建内容
	 * @param sheet
	 * @param excelSetting
	 */
	public void write(Sheet sheet, ExcelSetting excelSetting) {
		
	}
	
	/**
	 * 根据模板占位符替换内容
	 * @param sheet
	 * @param param
	 */
	public void write(Sheet sheet, Map<String, Object> param) {
		
	}

	/**
	 * 读取sheet
	 * @param <T>
	 * @param sheet
	 * @param clazz     自定义实体类
	 * @param rowNum    从第几行开始读取（从0开始计算）
	 * @return
	 * @throws Exception 
	 */
	public <T> List<T> read(Sheet sheet, Class<T> clazz, int rowNum) throws Exception {
		return null;
	}
	
	/**
	 * 设置下拉选项
	 * @param sheet
	 * @param colNum          第几个列（从0开始计算）
	 * @param startRow        第几个行设置开始（从0开始计算）
	 * @param endRow          第几个行设置结束（从0开始计算）
	 * @param selectDataArr   下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public void setSelect(Sheet sheet, int colIndex, int startRowIndex, int endRowIndex, String[] selectDataArr) {
		if (selectDataArr==null || selectDataArr.length==0) {
			return;
		}
		
		// 获取单元格样式
		CellStyle cellStyle = null;
		try {
			cellStyle = sheet.getRow(startRowIndex).getCell(colIndex).getCellStyle();
		} catch (Exception e) {
			// 如果没有该单元格存在，则使用默认的样式
			cellStyle = sheet.getWorkbook().createCellStyle();
		}
		// 自动换行
		cellStyle.setWrapText(true);
		
		Row row = null;
		Cell cell = null;
		
		for (int i=startRowIndex; i<=endRowIndex; i++) {
			row = sheet.getRow(i);
			if (row==null) {
				cell = sheet.createRow(i).createCell(colIndex);
			} else {
				cell = row.getCell(colIndex);
				if (cell==null) {
					cell = row.createCell(colIndex);
				}
			}
			
			cell.setCellStyle(cellStyle);
		}
		
		// 下拉的数据、起始行、终止行、起始列、终止列
		CellRangeAddressList addressList = new CellRangeAddressList(startRowIndex, endRowIndex, colIndex, colIndex);
		
		// 生成下拉框内容
		DataValidationHelper helper = sheet.getDataValidationHelper();
		DataValidationConstraint constraint = helper.createExplicitListConstraint(selectDataArr); 
		DataValidation dataValidation = helper.createValidation(constraint, addressList);
		
		// 设置数据有效性
		sheet.addValidationData(dataValidation);
	}

	/**
	 * 插入行
	 * @param sheet
	 * @param startRow
	 * @param rows
	 */
	public void insertRow(Sheet sheet, int startRow, int rows) {
		if (rows == 0) {
			return;
		}
		
		// 解决list占位符在最后一行时报错的BUG
		if ((startRow + 1) > sheet.getLastRowNum()) {
			sheet.createRow(startRow + 2);
		}
		
		/**
		 * startRow                  从下标为startRow的行开始移动
		 * endRow                    到下标为endRow的行结束移动
		 * n                         有多少行需要移动
		 * copyRowHeight             是否复制行高
		 * resetOriginalRowHeight    是否将原始行的高度设置为默认
		 */
		sheet.shiftRows(startRow + 1, sheet.getLastRowNum(), rows, true, false);
		
		RowHelper rowHelper = new RowHelper();
		
		for (int i=0; i<rows; i++) {
			Row sourceRow = null;
			Row targetRow = null;
			
			sourceRow = sheet.getRow(startRow);
			targetRow = sheet.createRow(++startRow);
			
			rowHelper.copyRow(sheet, sourceRow, targetRow);
		}
	}

	/**
	 * 设置只读
	 * @param sheet
	 * @param password
	 */
	public void setReadOnly(Sheet sheet, String password) {
		if (password==null || password.length()==0) {
			password = UUID.randomUUID().toString().replace("-", "");
		}
		
		sheet.protectSheet(password);
	}
	
	/**
	 * 设置下拉选项
	 * @param sheet
	 * @param colNum          第几个列（从0开始计算）
	 * @param startRow        第几个行设置开始（从0开始计算）
	 * @param endRow          第几个行设置结束（从0开始计算）
	 * @param selectDataArr   下拉数据，例如：new String[]{"2018", "2019", "2020"}
	 */
	public void setEditable(Sheet sheet, int[] colNumArr) {
		if (colNumArr==null || colNumArr.length==0) {
			return;
		}
		
		// 最大行
		int lastRowNum = sheet.getLastRowNum();
		for (int rowIndex=0; rowIndex<=lastRowNum; rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (row==null) {
				continue;
			}
			
			// 最大列
			int lastCellNum = (int) row.getLastCellNum();
			for (int colIndex=0; colIndex<lastCellNum; colIndex++) {
				// 判断该单元格是否要设置为只读
				if (ArrayHandler.contains(colNumArr, colIndex+1)) {
					continue;
				}
				
				// 获取单元格
				Cell cell = row.getCell(colIndex);
				if (cell==null) {
					cell = row.createCell(colIndex);
				}
				
				// 获取单元格样式
				CellStyle cellStyle = cell.getCellStyle();
				
				// 设置可编辑样式
				CellStyle cellUnlockedStyle = sheet.getWorkbook().createCellStyle();
				// 水平对齐方式
				cellUnlockedStyle.setAlignment(cellStyle.getAlignment());
				// 垂直对齐方式
				cellUnlockedStyle.setVerticalAlignment(cellStyle.getVerticalAlignment());
				// 自动换行
				cellUnlockedStyle.setWrapText(true);
				// 设置为可编辑的
				cellUnlockedStyle.setLocked(false);
				// 重新设置单元格样式
				cell.setCellStyle(cellUnlockedStyle);
			}
		}
	}

	/**
	 * 判断某单元格是否为合并单元格，左上角并返回合并行跨度，否则返回1
	 * @param sheet
	 * @param rowIdx
	 * @param colIdx
	 * @return
	 */
	public int getMergedRowSpanIfTopLeft(Sheet sheet, int rowIdx, int colIdx) {
		for (CellRangeAddress region : sheet.getMergedRegions()) {
			if (region.getFirstRow() == rowIdx && region.getFirstColumn() == colIdx) {
				// 是合并区的左上角
				return region.getLastRow() - region.getFirstRow() + 1;
			}
		}
		return 1; // 非合并区左上角，行数就是1
	}

	/**
	 * 复制Sheet
	 * @param workbook
	 * @param sourceSheetName    源Sheet名称
	 * @param targetSheetName    目标Sheet名称
	 * @param copyCount          复制次数
	 * @param printArea          打印区域
	 */
	public void copySheets(Workbook workbook, String sourceSheetName, String targetSheetName, int copyCount, SheetPrintArea printArea) {
		
	}

}
