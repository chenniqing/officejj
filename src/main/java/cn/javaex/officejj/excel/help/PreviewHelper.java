package cn.javaex.officejj.excel.help;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import cn.javaex.officejj.excel.ExcelUtils;

/**
 * 预览
 * 
 * @author 陈霓清
 */
public class PreviewHelper {
	
	private String IMAGE_FOLDER_PATH = "";
	
	/**
	 * Excel转Html
	 * @param filePath    excel文件路径，例如：D:\\Temp\\1.xlsx
	 * @return            返回生成的html文件的全路径，例如：D:\\Temp\\1_html\\1.html
	 * @throws Exception 
	 */
	public String excelToHtml(String filePath) throws Exception {
		File excelFile = new File(filePath);
		if (!excelFile.exists()) {
			throw new FileNotFoundException("指定文件不存在：" + filePath);
		}
		
		// 文件名称
		String excelName = excelFile.getName();
		// 文件后缀
		String fileSuffix = excelName.substring(excelName.lastIndexOf(".") + 1);
		// 生成的html文件名称
		String htmlName = excelName.replace("." + fileSuffix, ".html");
		String excelFolderPath = excelFile.getParent();
		String htmlFolderPath = excelFolderPath + File.separator + excelName.replace("." + fileSuffix, "") + "_html";
		
		// 判断html文件是否已存在
		File htmlFile = new File(htmlFolderPath + File.separator + htmlName);
		if (htmlFile.exists()) {
			return htmlFile.getAbsolutePath();
		} else {
			// 生成html文件上级文件夹
			File folder = new File(htmlFolderPath);
			if (!folder.exists()) {
				folder.mkdirs();
			}
		}
		
		// 图片目录
		this.IMAGE_FOLDER_PATH = htmlFolderPath + File.separator + "image";
		
		// 获取excel内容
		String excelHtmlContent = "";
		Workbook wb = ExcelUtils.getExcel(filePath);
		// v03
		if (wb instanceof HSSFWorkbook) {
			HSSFWorkbook hWb = (HSSFWorkbook) wb;
			excelHtmlContent = this.getExcelContent(hWb);
		}
		// v07
		else if (wb instanceof XSSFWorkbook) {
			XSSFWorkbook xWb = (XSSFWorkbook) wb;
			excelHtmlContent = this.getExcelContent(xWb);
		}
		
		// 向html文件中写入内容
		this.writeHtmlFile(htmlFile, excelHtmlContent);
		wb.close();
		
		// 返回html文件的绝对路径
		return htmlFile.getAbsolutePath();
	}

	/**
	 * 写入html文件
	 * @param htmlFile
	 * @param excelHtmlContent
	 */
	private void writeHtmlFile(File htmlFile, String excelHtmlContent) {
		OutputStream out = null;
		try {
			StringBuffer sb = new StringBuffer();
			sb.append("<!doctype html>");
			sb.append("<html>");
			sb.append("<head>");
			sb.append("<meta charset=\"utf-8\">");
			sb.append("<title>" + htmlFile.getName() + "</title>");
			sb.append("</head>");
			sb.append("<body>");
			sb.append(excelHtmlContent);
			sb.append("</body>");
			sb.append("</html>");
			
			out = new FileOutputStream(htmlFile);
			out.write(sb.toString().getBytes());
			out.flush();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(out);
		}
	}

	/**
	 * 获取Excel内容
	 * @param wb
	 * @param b
	 * @return
	 */
	private String getExcelContent(Workbook wb) {
		StringBuffer sb = new StringBuffer();
		
		// 获取每一个Sheet的内容
		for (int i=0; i<wb.getNumberOfSheets(); i++) {
			Sheet sheet = wb.getSheetAt(i);
			String sheetName = sheet.getSheetName();
			int lastRowNum = sheet.getLastRowNum();
			
			Map<String, String>[] map = this.getRowSpanColSpan(sheet);
			sb.append("<h3>"+sheetName+"</h3>");
			sb.append("<table style='border-collapse:collapse;' width='100%'>");
			
			// map等待存储excel图片
			Map<String, String> imageMap = new HashMap<String, String>();
			Map<String, PictureData> sheetImageMap = this.getSheetPictures(i, sheet, wb);
			if (sheetImageMap!=null && sheetImageMap.isEmpty()==false) {
				imageMap = this.printImage(sheetImageMap);
				this.printImageToWorkbook(imageMap, wb);
			}
			
			Row row = null;
			Cell cell = null;
			for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
				row = sheet.getRow(rowNum);
				if (row == null) {
					sb.append("<tr><td > &nbsp;</td></tr>");
					continue;
				}
				sb.append("<tr>");
				int lastColNum = row.getLastCellNum();
				for (int colNum = 0; colNum < lastColNum; colNum++) {
					cell = row.getCell(colNum);
					
					// 处理空白的单元格
					if (cell==null) {
						sb.append("<td>&nbsp;</td>");
						continue;
					}
					
					String imageHtml = "";
					String imageRowNum = i + "_" + rowNum + "_" + colNum;
					if (sheetImageMap != null && sheetImageMap.containsKey(imageRowNum)) {
						String imagePath = imageMap.get(imageRowNum);
						imageHtml = "<img src='" + imagePath + "' style='height:auto;'>";
					}
					String stringValue = ExcelUtils.getCellValue(cell);
					if (map[0].containsKey(rowNum + "," + colNum)) {
						String pointString = map[0].get(rowNum + "," + colNum);
						map[0].remove(rowNum + "," + colNum);
						int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
						int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
						int rowSpan = bottomeRow - rowNum + 1;
						int colSpan = bottomeCol - colNum + 1;
						sb.append("<td rowspan= '" + rowSpan + "' colspan= '" + colSpan + "' ");
					} else if (map[1].containsKey(rowNum + "," + colNum)) {
						map[1].remove(rowNum + "," + colNum);
						continue;
					} else {
						sb.append("<td ");
					}
					
					// 处理单元格样式
					this.handleExcelStyle(wb, sheet, cell, sb);
					
					sb.append(">");
					if (sheetImageMap != null && sheetImageMap.containsKey(imageRowNum)) {
						sb.append(imageHtml);
					}
					if (stringValue == null || "".equals(stringValue.trim())) {
						sb.append(" &nbsp; ");
					} else {
						// 将ascii码为160的空格转换为html下的空格（&nbsp;）
						sb.append(stringValue.replace(String.valueOf((char) 160), "&nbsp;"));
					}
					sb.append("</td>");
				}
				sb.append("</tr>");
			}
			
			sb.append("</table>");
		}
	
		return sb.toString();
	}

	/**
	 * 处理单元格样式
	 * @param wb
	 * @param sheet
	 * @param cell
	 * @param sb
	 */
	private void handleExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuffer sb) {
		CellStyle cellStyle = cell.getCellStyle();
		if (cellStyle != null) {
			HorizontalAlignment alignment = cellStyle.getAlignment();
			sb.append("align='" + this.convertAlignToHtml(alignment) + "' ");                     // 单元格内容的水平对齐方式
			VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignment();
			sb.append("valign='" + this.convertVerticalAlignToHtml(verticalAlignment) + "' ");    // 单元格内容的垂直对齐方式
			
			// v03
			if (wb instanceof HSSFWorkbook) {
				HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
				sb.append("style='");
				boolean bold  = hf.getBold();
				if (bold) {
					sb.append("font-weight:700;");                            // 字体加粗
				}
				short fontColor = hf.getColor();
				HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
				HSSFColor hc = palette.getColor(fontColor);
				sb.append("font-size: " + hf.getFontHeight() / 2 + "%;");     // 字体大小
				String fontColorStr = this.convertToStardColor(hc);
				if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
					sb.append("color:" + fontColorStr + ";");                 // 字体颜色
				}
				int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
				sb.append("width:" + columnWidth + "px;");
				short bgColor = cellStyle.getFillForegroundColor();
				hc = palette.getColor(bgColor);
				String bgColorStr = this.convertToStardColor(hc);
				if (bgColorStr!=null) {
					sb.append("background-color:" + bgColorStr + ";");        // 背景颜色
				}
				
				sb.append(getBorderStyle(palette, "border-top:", cellStyle.getTopBorderColor(), cellStyle.getTopBorderColor()));
				sb.append(getBorderStyle(palette, "border-right:", cellStyle.getRightBorderColor(), cellStyle.getRightBorderColor()));
				sb.append(getBorderStyle(palette, "border-bottom:", cellStyle.getBottomBorderColor(), cellStyle.getBottomBorderColor()));
				sb.append(getBorderStyle(palette, "border-left:", cellStyle.getLeftBorderColor(), cellStyle.getLeftBorderColor()));
			}
			// v07
			else if (wb instanceof XSSFWorkbook) {
				XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
				boolean bold = xf.getBold();
				sb.append("style='");
				if (bold) {
					sb.append("font-weight:700;");                            // 字体加粗
				}
				sb.append("font-size: " + xf.getFontHeight() / 2 + "%;");     // 字体大小
				int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
				sb.append("width:" + columnWidth + "px;");
				XSSFColor xssfColor = xf.getXSSFColor();
				if (xssfColor!=null) {
					String string = xssfColor.getARGBHex();
					if(string!=null && string.length()>0) {
						sb.append("color:#" + string.substring(2) + ";");     // 字体颜色
					}
				}
				
				XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
				if (bgColor!=null) {
					String argbHex = bgColor.getARGBHex();
					if(argbHex!=null && !"".equals(argbHex)) {
						sb.append("background-color:#" + argbHex.substring(2) + ";"); // 背景颜色
					}
				}
				
				sb.append(getBorderStyle("border-top:", cellStyle.getTopBorderColor(), ((XSSFCellStyle) cellStyle).getTopBorderXSSFColor()));
				sb.append(getBorderStyle("border-right:", cellStyle.getRightBorderColor(), ((XSSFCellStyle) cellStyle).getRightBorderXSSFColor()));
				sb.append(getBorderStyle("border-bottom:", cellStyle.getBottomBorderColor(), ((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor()));
				sb.append(getBorderStyle("border-left:", cellStyle.getLeftBorderColor(), ((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor()));
			}
			
			sb.append("' ");
		}
	}
	
	/**
	 * 颜色转化
	 * @param hc
	 * @return
	 */
	private String convertToStardColor(HSSFColor hc) {
		StringBuffer sb = new StringBuffer("");
		if (hc!=null) {
			if (hc.getIndex()==HSSFColor.HSSFColorPredefined.AUTOMATIC.getIndex()) {
				return null;
			}
			sb.append("#");
			for (int i=0; i<hc.getTriplet().length; i++) {
				sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
			}
		}
		
		return sb.toString();
	}

	/**
	 * 字符串前面填充0
	 * @param str
	 * @return
	 */
	private String fillWithZero(String str) {
		if (str!=null && str.length()<2) {
			 return "0" + str;
		}
		return str;
	}

	/**
	 * 获取边框样式
	 * @param pos
	 * @param s
	 * @param color
	 * @return
	 */
	private String getBorderStyle(HSSFPalette palette, String pos, short s, short color) {
		if (s==0) {
			return "";
		}
		
		String borderColorStr = convertToStardColor(palette.getColor(color));
		borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr;
		return pos + "solid " + borderColorStr + " 1px;";
	}
	
	/**
	 * 获取边框样式
	 * @param pos
	 * @param s
	 * @param color
	 * @return
	 */
	private String getBorderStyle(String pos, short s, XSSFColor color) {
		if (s==8) {
			return "";
		}
		
		if (color!=null) {
			String borderColorStr = color.getARGBHex();
			borderColorStr = (borderColorStr==null || borderColorStr.length()<1) ? "#000000" : borderColorStr.substring(2);
			return pos + "solid " + borderColorStr + " 1px;";
		}
		
		return "";
	}
	
	/**
	 * 单元格内容的垂直对齐方式
	 * @param verticalAlignment
	 * @return
	 */
	private String convertVerticalAlignToHtml(VerticalAlignment verticalAlignment) {
		String valign = "middle";
		
		switch (verticalAlignment) {
			case BOTTOM:
				valign = "bottom";
				break;
			case CENTER:
				valign = "center";
				break;
			case TOP:
				valign = "top";
				break;
			default:
				break;
		}
		
		return valign;
	}

	/**
	 * 单元格内容的水平对齐方式
	 * @param alignment
	 * @return
	 */
	private String convertAlignToHtml(HorizontalAlignment alignment) {
		String align = "left";
		
		switch (alignment) {
			case LEFT:
				align = "left";
				break;
			case CENTER:
				align = "center";
				break;
			case RIGHT:
				align = "right";
				break;
			default:
				break;
		}
		
		return align;
	}

	/**
	 * 对图片单元格赋值使其可读取到
	 * @param imageMap
	 * @param wb
	 */
	@SuppressWarnings("unused")
	private void printImageToWorkbook(Map<String, String> imageMap, Workbook wb) {
		Sheet sheet = null;
		Row row = null;
		Cell cell = null;
		
		String[] sheetRowCol = new String[3];
		for (String key : imageMap.keySet()) {
			sheetRowCol = key.split("_");
			sheet = wb.getSheetAt(Integer.parseInt(sheetRowCol[0]));
			row = sheet.getRow(Integer.parseInt(sheetRowCol[1]))==null ? sheet.createRow(Integer.parseInt(sheetRowCol[1])) : sheet.getRow(Integer.parseInt(sheetRowCol[1]));
			cell = row.getCell(Integer.parseInt(sheetRowCol[2]))==null ? row.createCell(Integer.parseInt(sheetRowCol[2])) : row.getCell(Integer.parseInt(sheetRowCol[2]));
		}
	}

	/**
	 * 输出图片
	 * @param map
	 * @return
	 */
	private Map<String, String> printImage(Map<String, PictureData> map) {
		Map<String, String> imageMap = new HashMap<String, String>();
		String imageName = null;
		try {
			Object key[] = map.keySet().toArray();
			for (int i=0; i<map.size(); i++) {
				// 获取图片流
				PictureData picture = map.get(key[i]);
				// 获取图片索引
				String pictureName = key[i].toString();
				// 获取图片格式
				String ext = picture.suggestFileExtension();
				
				byte[] data = picture.getData();
				File uploadFile = new File(this.IMAGE_FOLDER_PATH);
				if (!uploadFile.exists()) {
					uploadFile.mkdirs();
				}
				imageName = pictureName + "-" + UUID.randomUUID().toString().replace("-", "") + "." + ext;
				FileOutputStream out = new FileOutputStream(this.IMAGE_FOLDER_PATH + File.separator + imageName);
				imageMap.put(pictureName, this.IMAGE_FOLDER_PATH + File.separator + imageName);
				out.write(data);
				out.flush();
				out.close();
			}
		} catch (Exception e) {
			
		}
		
		return imageMap;
	}

	/**
	 * 获取Excel图片
	 * @param sheetNum
	 * @param sheet
	 * @param wb
	 * @return
	 */
	private Map<String, PictureData> getSheetPictures(int sheetNum, Sheet sheet, Workbook wb) {
		// v03
		if (wb instanceof HSSFWorkbook) {
			return getSheetPicturesFromXls(sheetNum, (HSSFSheet) sheet, (HSSFWorkbook) wb);
		}
		// v07
		else if (wb instanceof XSSFWorkbook) {
			return getSheetPicturesFromXlsx(sheetNum, (XSSFSheet) sheet);
			
		}
		return null;
		
	}

	/**
	 * 获取xls图片
	 * @param sheetNum
	 * @param sheet
	 * @param wb
	 * @return
	 */
	private Map<String, PictureData> getSheetPicturesFromXls(int sheetNum, HSSFSheet sheet, HSSFWorkbook wb) {
		Map<String, PictureData> sheetImageMap = new HashMap<String, PictureData>();
		List<HSSFPictureData> pictures = wb.getAllPictures();
		if (pictures==null || pictures.isEmpty()) {
			return null;
		}
		
		if (sheet.getDrawingPatriarch()==null) {
			return null;
		}
		
		for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
			HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
			shape.getLineWidth();
			if (shape instanceof HSSFPicture) {
				HSSFPicture pic = (HSSFPicture) shape;
				int pictureIndex = pic.getPictureIndex() - 1;
				HSSFPictureData picData = pictures.get(pictureIndex);
				String picIndex = String.valueOf(sheetNum) + "_" + String.valueOf(anchor.getRow1()) + "_" + String.valueOf(anchor.getCol1());
				sheetImageMap.put(picIndex, picData);
			}
		}
		
		return sheetImageMap;
	}

	/**
	 * 获取xlsx图片
	 * @param sheetNum
	 * @param sheet
	 * @return
	 */
	private Map<String, PictureData> getSheetPicturesFromXlsx(int sheetNum, XSSFSheet sheet) {
		Map<String, PictureData> sheetImageMap = new HashMap<String, PictureData>();
		
		for (POIXMLDocumentPart dr : sheet.getRelations()) {
			if (dr instanceof XSSFDrawing) {
				XSSFDrawing drawing = (XSSFDrawing) dr;
				List<XSSFShape> shapes = drawing.getShapes();
				for (XSSFShape shape : shapes) {
					XSSFPicture pic = (XSSFPicture) shape;
					XSSFClientAnchor anchor = pic.getPreferredSize();
					CTMarker ctMarker = anchor.getFrom();
					String picIndex = String.valueOf(sheetNum) + "_" + ctMarker.getRow() + "_" + ctMarker.getCol();
					sheetImageMap.put(picIndex, pic.getPictureData());
				}
			}
		}
		
		return sheetImageMap;
	}

	/**
	 * 获取行、列
	 * @param sheet
	 * @return
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	private Map<String, String>[] getRowSpanColSpan(Sheet sheet) {
		Map<String, String> map0 = new HashMap<String, String>();
		Map<String, String> map1 = new HashMap<String, String>();
		int mergedNum = sheet.getNumMergedRegions();
		CellRangeAddress range = null;
		
		for (int i=0; i<mergedNum; i++) {
			range = sheet.getMergedRegion(i);
			int topRow = range.getFirstRow();
			int topCol = range.getFirstColumn();
			int bottomRow = range.getLastRow();
			int bottomCol = range.getLastColumn();
			
			map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
			
			int tempRow = topRow;
			while (tempRow<=bottomRow) {
				int tempCol = topCol;
				while (tempCol<=bottomCol) {
					map1.put(tempRow + "," + tempCol, "");
					tempCol++;
				}
				tempRow++;
			}
			
			map1.remove(topRow + "," + topCol);
		}
		
		Map[] map = {map0, map1};
		return map;
	}
	
}
