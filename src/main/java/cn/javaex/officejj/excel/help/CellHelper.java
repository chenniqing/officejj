package cn.javaex.officejj.excel.help;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
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
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;

import org.apache.poi.util.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.springframework.beans.BeanUtils;

import cn.javaex.officejj.common.entity.Font;
import cn.javaex.officejj.common.entity.Picture;
import cn.javaex.officejj.common.entity.RGB;
import cn.javaex.officejj.common.util.ImageHandler;
import cn.javaex.officejj.excel.style.ICellStyle;

/**
 * Cell
 * 
 * @author 陈霓清
 */
public class CellHelper extends SheetHelper {
	
	/**
	 * 自动获得它所处的合并区域
	 * @param cell
	 * @return
	 */
	public CellRangeAddress getMergedRegion(Cell cell) {
		Sheet sheet = cell.getSheet();
	    int rowIndex = cell.getRowIndex();
	    int colIndex = cell.getColumnIndex();
	    int mergedCount = sheet.getNumMergedRegions();
	    for (int i = 0; i < mergedCount; i++) {
	        CellRangeAddress range = sheet.getMergedRegion(i);
	        if (range.isInRange(rowIndex, colIndex)) {
	            return range;
	        }
	    }
	    return null; // 说明该cell不在合并区域内
	}
	
	/**
	 * 提取${xx}中的文本
	 * @param str
	 * @return
	 */
	public List<String> getPlaceholders(String str) {
		List<String> list = new ArrayList<String>();
		
		String patern = "(?<=\\$\\{)[^\\}]+";
		Pattern pattern = Pattern.compile(patern);
		Matcher matcher = pattern.matcher(str);
		while (matcher.find()) {
			list.add(matcher.group());
		}
		
		return list;
	}
	
	/**
	 * 设置图片
	 * @param cell
	 * @param imgStream
	 * @param fileSuffix
	 * @param width
	 * @param height
	 */
	public void setImage(Cell cell, InputStream imgStream, String fileSuffix, Integer width, Integer height) {
		if (imgStream == null) {
			cell.setCellValue("");
			return;
		}
		
		Sheet sheet = cell.getSheet();
		
		if (width != null && height != null) {
			sheet.setColumnWidth(cell.getColumnIndex(), width * 32);
			cell.getRow().setHeight((short) (height * 15));
		}
		
		try {
			ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
			
			// 使用传入的InputStream和图片后缀名
			fileSuffix = fileSuffix.toLowerCase();
	        BufferedImage bufferImg = ImageIO.read(imgStream);
	        ImageIO.write(bufferImg, fileSuffix, byteArrayOut);
	        
	        Drawing<?> patriarch = sheet.createDrawingPatriarch();
	        
	        int imageType = Workbook.PICTURE_TYPE_JPEG;
			if ("png".equals(fileSuffix)) {
				imageType = Workbook.PICTURE_TYPE_PNG;
			}
			
			ClientAnchor anchor = null;
			CellRangeAddress region = getMergedRegion(cell);
			int marginX = 0;
			int marginY = 0;
			
			if (region != null) {
				anchor = new XSSFClientAnchor(
						marginX, marginY,
						-marginX, -marginY,
						region.getFirstColumn(), region.getFirstRow(),
						region.getLastColumn() + 1, region.getLastRow() + 1
				);
			} else {
				anchor = new XSSFClientAnchor(
						marginX, marginY,
						-marginX, -marginY,
						(short) cell.getColumnIndex(), cell.getRow().getRowNum(),
						(short) (cell.getColumnIndex() + 1), cell.getRow().getRowNum() + 1
				);
			}
			
			anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
			patriarch.createPicture(anchor, sheet.getWorkbook().addPicture(byteArrayOut.toByteArray(), imageType));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 设置图片
	 * @param cell
	 * @param path
	 * @throws IOException 
	 */
	public void setImage(Cell cell, String path, Integer width, Integer height) {
	    if (path == null || path.length() == 0) {
	        cell.setCellValue("");
	        return;
	    }
	 
	    Sheet sheet = cell.getSheet();
	 
	    if (width != null && height != null) {
	        sheet.setColumnWidth(cell.getColumnIndex(), width * 32);
	        cell.getRow().setHeight((short) (height * 15));
	    }
	 
	    try {
	        // 图片后缀
	        String fileSuffix = path.substring(path.lastIndexOf(".") + 1).toLowerCase();
	        if (fileSuffix.length() > 5) {
	            fileSuffix = "jpg";
	        }
	 
	        // 读取图片字节数据
	        InputStream imageStream = ImageHandler.getImageStream(path);
	        byte[] bytes = IOUtils.toByteArray(imageStream);
	        imageStream.close();
	 
	        // 设定内边距（px）
	        int paddingPx = 20;
	        int paddingEMU = paddingPx * 9525;
	 
	        Drawing<?> patriarch = sheet.createDrawingPatriarch();
	 
	        int imageType = Workbook.PICTURE_TYPE_JPEG;
	        if ("png".equals(fileSuffix)) {
	            imageType = Workbook.PICTURE_TYPE_PNG;
	        }
	 
	        ClientAnchor anchor = null;
	        // 判断是否是合并单元格
	        CellRangeAddress region = getMergedRegion(cell);
	 
	        if (region != null) {
	            anchor = new XSSFClientAnchor(
	                    paddingEMU, paddingEMU, // dx1, dy1: 左上内边距
	                    -paddingEMU, -paddingEMU, // dx2, dy2: 右下内边距(负值内缩)
	                    region.getFirstColumn(), region.getFirstRow(), // 起始单元格
	                    region.getLastColumn() + 1, region.getLastRow() + 1 // 结束单元格（闭区间+1）
	            );
	        } else {
	            anchor = new XSSFClientAnchor(
	                    paddingEMU, paddingEMU,
	                    -paddingEMU, -paddingEMU,
	                    (short) cell.getColumnIndex(), cell.getRow().getRowNum(),
	                    (short) (cell.getColumnIndex() + 1), cell.getRow().getRowNum() + 1
	            );
	        }
	 
	        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
	 
	        int picIndex = sheet.getWorkbook().addPicture(bytes, imageType);
	        patriarch.createPicture(anchor, picIndex);
	 
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	}
	
	/**
	 * 替换模板中的占位符（占位符独占一格单元格）
	 * @param cell
	 * @param obj
	 */
	public void setValue(Cell cell, Object obj) {
		if (obj==null) {
			cell.setCellValue((String) obj);
		}
		else if (obj instanceof String) {
			if (obj.equals(PLACEHOLDER_CLEAR) == false) {
				cell.setCellValue((String) obj);
			}
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
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
			cell.setCellValue(dtf.format((LocalDateTime) obj));
		}
		else if (obj instanceof LocalDate) {
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd");
			cell.setCellValue(dtf.format((LocalDate) obj));
		}
		else if (obj instanceof Date) {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
			cell.setCellValue(sdf.format((Date) obj));
		}
		// 自定义字体
		else if (obj instanceof Font) {
			Font font = (Font) obj;
			
			XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getSheet().getWorkbook().createCellStyle();
			BeanUtils.copyProperties((XSSFCellStyle) cell.getCellStyle(), cellStyle);
			cellStyle.setFont(setXSSFFont(cell, font));
			
			cell.setCellValue(font.getText());
			cell.setCellStyle(cellStyle);
		}
		// 自定义图片
		else if (obj instanceof Picture) {
			cell.setCellValue("");
			Picture picture = (Picture) obj;
			if (picture.getUrl() != null && !picture.getUrl().isEmpty()) {
				this.setImage(cell, picture.getUrl(), picture.getWidth(), picture.getHeight());
			} else if (picture.getData() != null) {
				try (InputStream is = new ByteArrayInputStream(picture.getData())) {
					this.setImage(cell, is, picture.getFileSuffix(), picture.getWidth(), picture.getHeight());
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		else {
			cell.setCellValue(obj.toString());
		}
	}

	/**
	 * 替换模板中的占位符（占位符共享格单元格）
	 * @param cell
	 * @param placeholders    占位符集合
	 * @param param           传入的替换参数
	 */
	public void setValue(Cell cell, List<String> placeholders, Map<String, Object> param) {
		// 定义一个Map，用来存储单元格中需要添加字体样式的关键字
		Map<String, Font> map = new HashMap<String, Font>();
		
		// 获取单元格中原本的内容
		String cellValue = cell.getRichStringCellValue().getString();
		
		// 遍历替换单元格中的占位符集合
		for (String placeholder : placeholders) {
			Object obj = param.get(placeholder);
			if (obj==null) {
				continue;
			}
			
			if (obj instanceof Font) {
				// 自定义字体样式
				Font font = (Font) obj;
				map.put(font.getText(), font);    // 存储单元格中需要添加字体样式的关键字
				cellValue = cellValue.replace("${" + placeholder + "}", font.getText());
			} else {
				// 其他情况都认为是字符串
				cellValue = cellValue.replace("${" + placeholder + "}", obj.toString());
			}
		}
		
		if (map.isEmpty()) {
			cell.setCellValue(cellValue);
		} else {
			// 为指定的文本设置字体样式
			XSSFRichTextString richString = new XSSFRichTextString(cellValue);
			
			for (Map.Entry<String, Font> entry : map.entrySet()) {
				String key = entry.getKey();
				XSSFFont fontSetting = setXSSFFont(cell, entry.getValue());
				richString.applyFont(cellValue.indexOf(key), cellValue.indexOf(key) + key.length(), fontSetting);
			}
			
			cell.setCellValue(richString);
		}
	}
	
	/**
	 * 设置字体样式
	 * @param cell
	 * @param font
	 * @return
	 */
	private XSSFFont setXSSFFont(Cell cell, Font font) {
		XSSFFont fontSetting = (XSSFFont) cell.getSheet().getWorkbook().createFont();
		fontSetting.setFontName(font.getFontFamily());
		fontSetting.setBold(font.getBold());
		fontSetting.setItalic(font.getItalic());
		fontSetting.setStrikeout(font.getStrike());
		if (font.getFontSize()!=null) {
			fontSetting.setFontHeightInPoints((short) font.getFontSize().intValue());
		}
		if (font.getColor()!=null) {
			RGB rgb = new RGB(font.getColor());
			XSSFColor color = new XSSFColor(new java.awt.Color(rgb.getRed(), rgb.getGreen(), rgb.getBlue()), new DefaultIndexedColorMap());
			fontSetting.setColor(color);
		}
		
		return fontSetting;
	}

	/**
	 * 设置单元格样式
	 * @param cell
	 * @param clazz
	 */
	public void setCellStyle(Cell cell, Class<?> clazz) {
		try {
			Sheet sheet = cell.getSheet();
			
			ICellStyle styleProvider = (ICellStyle) clazz.getDeclaredConstructor().newInstance();
			cell.setCellStyle(styleProvider.createDataStyle(sheet.getWorkbook()));
		} catch (Exception e) {
			throw new RuntimeException("设置单元格样式失败", e);
		}
	}

}
