package cn.javaex.officejj.excel.help;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Workbook;
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

/**
 * Cell
 * 
 * @author 陈霓清
 */
public class CellHelper extends SheetHelper {
	
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
	 * @param path
	 * @throws IOException 
	 */
	public void setImage(Cell cell, String path, Integer width, Integer height) {
		try {
			if (path==null || path.length()==0) {
				cell.setCellValue("");
				return;
			}
			
			if (width!=null) {
				cell.getSheet().setColumnWidth(cell.getColumnIndex(), width * 32);
			}
			if (height==null) {
				height = 100;
			}
			cell.getRow().setHeight((short) (height * 15));
			
			ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
			// 图片后缀
			String fileSuffix = path.substring(path.lastIndexOf(".") + 1).toLowerCase();
			if (fileSuffix.length()>5) {
				fileSuffix = "jpg";
			}
			
			BufferedImage bufferImg = ImageIO.read(ImageHandler.getImageStream(path));
			ImageIO.write(bufferImg, fileSuffix, byteArrayOut);
			
			Drawing<?> patriarch = cell.getSheet().createDrawingPatriarch();
			ClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, (short) cell.getColumnIndex(), cell.getRow().getRowNum(), (short) (cell.getColumnIndex() + 1), cell.getRow().getRowNum() + 1);
			
			int imageType = Workbook.PICTURE_TYPE_JPEG;
			if ("png".equals(fileSuffix)) {
				imageType = Workbook.PICTURE_TYPE_PNG;
			}
			patriarch.createPicture(anchor, cell.getSheet().getWorkbook().addPicture(byteArrayOut.toByteArray(), imageType));
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
			return;
		}
		else if (obj instanceof String) {
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
			this.setImage(cell, picture.getUrl(), picture.getWidth(), picture.getHeight());
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

}
