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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import cn.javaex.officejj.common.entity.Font;
import cn.javaex.officejj.common.entity.Picture;
import cn.javaex.officejj.common.util.PropertyHandler;
import cn.javaex.officejj.common.util.ImageHandler;
import cn.javaex.officejj.excel.style.ICellStyle;

/**
 * Cell
 *
 * @author 陈霓清
 */
public class CellHelper extends SheetHelper {

	// POI 字体缓存：同样属性只建一次
	private final Map<String, org.apache.poi.ss.usermodel.Font> poiFontCache = new HashMap<>();
	// 样式缓存：baseStyle + fontKey => 新样式
	private final Map<String, CellStyle> styleCache = new HashMap<>();
	// Excel图片默认展示宽高，单位：像素
	private static final int DEFAULT_IMAGE_WIDTH = 120;
	private static final int DEFAULT_IMAGE_HEIGHT = 80;
	private static final int DEFAULT_IMAGE_PADDING = 5;
	private static final int EXCEL_COLUMN_WIDTH_CORRECTION_PIXELS = 14;

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
		if (str==null || str.length()==0) {
			return list;
		}

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

		try {
			byte[] bytes = IOUtils.toByteArray(imgStream);
			this.setImage(cell, bytes, fileSuffix, width, height, true, true, DEFAULT_IMAGE_PADDING);
		} catch (Exception e) {
			throw new RuntimeException("设置Excel图片失败", e);
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

	    try {
			// 读取图片字节数据
			byte[] bytes = null;
			try (InputStream imageStream = ImageHandler.getImageStream(path)) {
				bytes = IOUtils.toByteArray(imageStream);
			}

			this.setImage(cell, bytes, this.getImageSuffix(path), width, height, true, true, DEFAULT_IMAGE_PADDING);
	    } catch (Exception e) {
	        throw new RuntimeException("设置Excel图片失败：" + path, e);
	    }
	}

	/**
	 * 设置图片，支持自动撑开单元格、等比缩放和居中展示。
	 * @param cell
	 * @param picture
	 */
	public void setImage(Cell cell, Picture picture) {
		if (picture==null) {
			cell.setCellValue("");
			return;
		}

		try {
			ImageData imageData = this.getImageData(picture);
			if (imageData==null) {
				cell.setCellValue("");
				return;
			}

			this.setImage(cell, imageData.getBytes(), imageData.getFileSuffix(),
					picture.getWidth(), picture.getHeight(),
					!Boolean.FALSE.equals(picture.getKeepRatio()),
					!Boolean.FALSE.equals(picture.getResizeCell()),
					picture.getEffectivePadding());
		} catch (Exception e) {
			throw new RuntimeException("设置Excel图片失败", e);
		}
	}

	/**
	 * 在一个单元格里写入多张图片，默认按网格排版。
	 * @param cell
	 * @param imageList 图片集合，元素支持 Picture 或图片路径字符串
	 */
	public void setImages(Cell cell, List<?> imageList) {
		this.setImages(cell, imageList, null, null);
	}

	/**
	 * 在一个单元格里写入多张图片，默认按网格排版。
	 * @param cell
	 * @param imageList 图片集合，元素支持 Picture 或图片路径字符串
	 * @param width 单张图片最大展示宽度，允许为空
	 * @param height 单张图片最大展示高度，允许为空
	 */
	public void setImages(Cell cell, List<?> imageList, Integer width, Integer height) {
		List<Picture> pictureList = this.toPictureList(imageList, width, height);
		if (pictureList.isEmpty()) {
			cell.setCellValue("");
			return;
		}

		int slotWidth = this.getMaxImageWidth(pictureList);
		int slotHeight = this.getMaxImageHeight(pictureList);
		int maxColumns = pictureList.get(0).getEffectiveMaxColumns();
		int columns = Math.min(maxColumns, pictureList.size());
		int rows = (pictureList.size() + columns - 1) / columns;
		boolean resizeCell = !Boolean.FALSE.equals(pictureList.get(0).getResizeCell());
		if (resizeCell) {
			this.resizeImageCell(cell, slotWidth * columns, slotHeight * rows);
		}

		for (int i=0; i<pictureList.size(); i++) {
			Picture picture = pictureList.get(i);
			try {
				ImageData imageData = this.getImageData(picture);
				if (imageData==null) {
					continue;
				}

				int colIndex = i % columns;
				int rowIndex = i / columns;
				this.drawImage(cell, imageData.getBytes(), imageData.getFileSuffix(),
						slotWidth, slotHeight,
						!Boolean.FALSE.equals(picture.getKeepRatio()),
						picture.getEffectivePadding(),
						colIndex * slotWidth,
						rowIndex * slotHeight,
						false);
			} catch (Exception e) {
				throw new RuntimeException("设置Excel图片失败", e);
			}
		}
	}

	/**
	 * 写入图片字节，并根据原图比例计算最终展示尺寸。
	 * @param cell
	 * @param imageBytes
	 * @param fileSuffix
	 * @param maxWidth
	 * @param maxHeight
	 * @param keepRatio
	 * @param resizeCell
	 * @param padding
	 */
	private void setImage(Cell cell, byte[] imageBytes, String fileSuffix, Integer maxWidth, Integer maxHeight,
			boolean keepRatio, boolean resizeCell, Integer padding) {
		if (imageBytes==null || imageBytes.length==0) {
			cell.setCellValue("");
			return;
		}

		fileSuffix = this.normalizeImageSuffix(fileSuffix);
		padding = padding==null || padding<0 ? DEFAULT_IMAGE_PADDING : padding;
		boolean fillCell = (maxWidth==null || maxWidth<=0) && (maxHeight==null || maxHeight<=0);
		int imageWidth = maxWidth==null || maxWidth<=0 ? DEFAULT_IMAGE_WIDTH : maxWidth;
		int imageHeight = maxHeight==null || maxHeight<=0 ? DEFAULT_IMAGE_HEIGHT : maxHeight;
		int boxWidth = imageWidth + padding * 2;
		int boxHeight = imageHeight + padding * 2;

		try {
			if (fillCell) {
				// 未指定图片宽高时，沿用模板单元格或合并区域的现有尺寸作为展示盒子，并在盒子内保留内边距。
				CellRangeAddress mergedRegion = this.getMergedRegion(cell);
				if (mergedRegion!=null) {
					boxWidth = this.getMergedRegionWidthInPixels(cell.getSheet(), mergedRegion);
					boxHeight = this.getMergedRegionHeightInPixels(cell.getSheet(), mergedRegion);
				} else {
					boxWidth = this.getColumnWidthInPixels(cell.getSheet(), cell.getColumnIndex());
					boxHeight = this.getRowHeightInPixels(cell.getSheet(), cell.getRowIndex());
				}
			} else if (resizeCell) {
				this.resizeImageCell(cell, boxWidth, boxHeight);
			}
			this.drawImage(cell, imageBytes, fileSuffix, boxWidth, boxHeight, keepRatio, padding, 0, 0, fillCell);
		} catch (Exception e) {
			throw new RuntimeException("设置Excel图片失败", e);
		}
	}

	/**
	 * 在单元格里的指定槽位绘制图片。
	 * @param cell
	 * @param imageBytes
	 * @param fileSuffix
	 * @param boxWidth
	 * @param boxHeight
	 * @param keepRatio
	 * @param padding
	 * @param offsetX 槽位左上角X偏移，单位：像素
	 * @param offsetY 槽位左上角Y偏移，单位：像素
	 * @throws IOException
	 */
	private void drawImage(Cell cell, byte[] imageBytes, String fileSuffix, int boxWidth, int boxHeight,
			boolean keepRatio, int padding, int offsetX, int offsetY, boolean useMergedRegion) throws IOException {
		BufferedImage bufferImg = ImageIO.read(new ByteArrayInputStream(imageBytes));
		if (bufferImg==null) {
			throw new IllegalArgumentException("图片流不是有效图片");
		}

		fileSuffix = this.normalizeImageSuffix(fileSuffix);
		int contentWidth = Math.max(1, boxWidth - padding * 2);
		int contentHeight = Math.max(1, boxHeight - padding * 2);
		ImageSize imageSize = useMergedRegion
				? new ImageSize(contentWidth, contentHeight)
				: this.calculateImageSize(bufferImg.getWidth(), bufferImg.getHeight(), contentWidth, contentHeight, keepRatio);
		byte[] pictureBytes = this.toPictureBytes(bufferImg, fileSuffix);
		int imageType = this.getPoiImageType(fileSuffix);
		ClientAnchor anchor = this.createImageAnchor(cell, boxWidth, boxHeight, imageSize, padding, offsetX, offsetY, useMergedRegion);

		Drawing<?> patriarch = cell.getSheet().createDrawingPatriarch();
		int picIndex = cell.getSheet().getWorkbook().addPicture(pictureBytes, imageType);
		patriarch.createPicture(anchor, picIndex);
	}

	/**
	 * 按最大宽高计算图片展示尺寸。
	 * @param originWidth
	 * @param originHeight
	 * @param maxWidth
	 * @param maxHeight
	 * @param keepRatio
	 * @return
	 */
	private ImageSize calculateImageSize(int originWidth, int originHeight, int maxWidth, int maxHeight, boolean keepRatio) {
		if (!keepRatio || originWidth<=0 || originHeight<=0) {
			return new ImageSize(maxWidth, maxHeight);
		}

		double ratio = Math.min(maxWidth * 1.0D / originWidth, maxHeight * 1.0D / originHeight);
		if (ratio<=0) {
			ratio = 1.0D;
		}

		int width = Math.max(1, (int) Math.round(originWidth * ratio));
		int height = Math.max(1, (int) Math.round(originHeight * ratio));
		return new ImageSize(width, height);
	}

	/**
	 * 自动调整图片所在单元格大小。
	 * @param cell
	 * @param width
	 * @param height
	 */
	private void resizeImageCell(Cell cell, int width, int height) {
		Sheet sheet = cell.getSheet();
		// 指定图片尺寸时只在空间不足时撑大模板，避免把用户已经设置好的较大列宽/行高缩小。
		if (this.getColumnWidthInPixels(sheet, cell.getColumnIndex())<width) {
			sheet.setColumnWidth(cell.getColumnIndex(), this.pixelToColumnWidth(width));
		}
		if (this.getRowHeightInPixels(sheet, cell.getRowIndex())<height) {
			cell.getRow().setHeight((short) (height * 15));
		}
	}

	/**
	 * 把像素宽度换算成 Excel 列宽单位。
	 * Excel 的列宽不是像素，直接用固定倍数会偏小，导致图片右侧看起来没有内边距。
	 * Excel/WPS 渲染列宽时会额外带上边界、网格线和字体度量补偿；POI 的 getColumnWidthInPixels 读回值偏小，
	 * 但客户端实际显示会更宽。换算时扣掉这段补偿，避免自动撑开的图片列右侧出现明显多余留白。
	 * @param pixels 目标像素宽度
	 * @return Excel 列宽单位，1 个字符宽度等于 256 个单位
	 */
	private int pixelToColumnWidth(int pixels) {
		if (pixels<=0) {
			return 1;
		}

		double visualPixels = Math.max(1.0D, pixels - EXCEL_COLUMN_WIDTH_CORRECTION_PIXELS);
		double characterWidth = visualPixels / Units.DEFAULT_CHARACTER_WIDTH;
		int columnWidth = (int) Math.ceil(characterWidth * 256.0D);
		return Math.max(1, Math.min(255 * 256, columnWidth));
	}

	/**
	 * 创建居中的图片锚点。
	 * @param cell
	 * @param boxWidth
	 * @param boxHeight
	 * @param imageSize
	 * @param padding
	 * @return
	 */
	private ClientAnchor createImageAnchor(Cell cell, int boxWidth, int boxHeight, ImageSize imageSize, int padding,
			int offsetX, int offsetY, boolean useMergedRegion) {
		int col1 = cell.getColumnIndex();
		int row1 = cell.getRowIndex();

		int contentWidth = Math.max(1, boxWidth - padding * 2);
		int contentHeight = Math.max(1, boxHeight - padding * 2);
		int left = offsetX + padding + Math.max(0, (contentWidth - imageSize.getWidth()) / 2);
		int top = offsetY + padding + Math.max(0, (contentHeight - imageSize.getHeight()) / 2);
		int right = left + imageSize.getWidth();
		int bottom = top + imageSize.getHeight();

		if (useMergedRegion) {
			CellRangeAddress mergedRegion = this.getMergedRegion(cell);
			int firstColumn = mergedRegion==null ? col1 : mergedRegion.getFirstColumn();
			int firstRow = mergedRegion==null ? row1 : mergedRegion.getFirstRow();
			int lastColumn = mergedRegion==null ? col1 : mergedRegion.getLastColumn();
			int lastRow = mergedRegion==null ? row1 : mergedRegion.getLastRow();
			int paddingEmu = Units.pixelToEMU(padding);
			XSSFClientAnchor anchor = new XSSFClientAnchor(
					paddingEmu,
					paddingEmu,
					-paddingEmu,
					-paddingEmu,
					firstColumn, firstRow,
					lastColumn + 1, lastRow + 1);
			anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
			return anchor;
		}

		XSSFClientAnchor anchor = new XSSFClientAnchor(
				Units.pixelToEMU(left),
				Units.pixelToEMU(top),
				Units.pixelToEMU(right),
				Units.pixelToEMU(bottom),
				col1, row1, col1, row1);
		anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
		return anchor;
	}

	/**
	 * 计算合并区域的总宽度，单位：像素。
	 * @param sheet
	 * @param region
	 * @return
	 */
	private int getMergedRegionWidthInPixels(Sheet sheet, CellRangeAddress region) {
		int width = 0;
		for (int colIndex=region.getFirstColumn(); colIndex<=region.getLastColumn(); colIndex++) {
			width += this.getColumnWidthInPixels(sheet, colIndex);
		}

		return Math.max(1, width);
	}

	/**
	 * 计算合并区域的总高度，单位：像素。
	 * @param sheet
	 * @param region
	 * @return
	 */
	private int getMergedRegionHeightInPixels(Sheet sheet, CellRangeAddress region) {
		int height = 0;
		for (int rowIndex=region.getFirstRow(); rowIndex<=region.getLastRow(); rowIndex++) {
			height += this.getRowHeightInPixels(sheet, rowIndex);
		}

		return Math.max(1, height);
	}

	/**
	 * 获取列宽像素值，最小返回 1，避免异常模板导致锚点计算为 0。
	 * @param sheet
	 * @param colIndex
	 * @return
	 */
	private int getColumnWidthInPixels(Sheet sheet, int colIndex) {
		return Math.max(1, (int) sheet.getColumnWidthInPixels(colIndex));
	}

	/**
	 * 获取行高像素值，兼容未显式创建的行。
	 * @param sheet
	 * @param rowIndex
	 * @return
	 */
	private int getRowHeightInPixels(Sheet sheet, int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		float heightInPoints = row==null ? sheet.getDefaultRowHeightInPoints() : row.getHeightInPoints();
		return Math.max(1, Math.round(heightInPoints * Units.PIXEL_DPI / Units.POINT_DPI));
	}

	/**
	 * 把外部传入的图片集合统一转换为 Picture 集合。
	 * @param imageList
	 * @param width
	 * @param height
	 * @return
	 */
	private List<Picture> toPictureList(List<?> imageList, Integer width, Integer height) {
		List<Picture> pictureList = new ArrayList<Picture>();
		if (imageList==null || imageList.isEmpty()) {
			return pictureList;
		}

		for (Object obj : imageList) {
			Picture picture = null;
			if (obj instanceof Picture) {
				picture = (Picture) obj;
			}
			else if (obj instanceof String) {
				picture = new Picture((String) obj);
			}
			else if (obj instanceof byte[]) {
				picture = new Picture();
				picture.setData((byte[]) obj);
			}

			if (picture==null) {
				continue;
			}
			if (width!=null) {
				picture.setWidth(width);
			}
			if (height!=null) {
				picture.setHeight(height);
			}
			pictureList.add(picture);
		}

		return pictureList;
	}

	/**
	 * 读取 Picture 的图片内容。
	 * @param picture
	 * @return
	 * @throws IOException
	 */
	private ImageData getImageData(Picture picture) throws IOException {
		if (picture==null) {
			return null;
		}
		if (picture.getData()!=null && picture.getData().length>0) {
			return new ImageData(picture.getData(), picture.getFileSuffix());
		}
		if (picture.getUrl()!=null && picture.getUrl().length()>0) {
			try (InputStream imageStream = ImageHandler.getImageStream(picture.getUrl())) {
				return new ImageData(IOUtils.toByteArray(imageStream), picture.getFileSuffix()==null ? this.getImageSuffix(picture.getUrl()) : picture.getFileSuffix());
			}
		}

		return null;
	}

	/**
	 * 取得多图布局里单个槽位的最大宽度。
	 * @param pictureList
	 * @return
	 */
	private int getMaxImageWidth(List<Picture> pictureList) {
		int width = DEFAULT_IMAGE_WIDTH;
		for (Picture picture : pictureList) {
			width = Math.max(width, picture.getEffectiveWidth() + picture.getEffectivePadding() * 2);
		}
		return width;
	}

	/**
	 * 取得多图布局里单个槽位的最大高度。
	 * @param pictureList
	 * @return
	 */
	private int getMaxImageHeight(List<Picture> pictureList) {
		int height = DEFAULT_IMAGE_HEIGHT;
		for (Picture picture : pictureList) {
			height = Math.max(height, picture.getEffectiveHeight() + picture.getEffectivePadding() * 2);
		}
		return height;
	}

	/**
	 * 把图片按目标后缀重新写成POI可识别的字节数组。
	 * @param bufferImg
	 * @param fileSuffix
	 * @return
	 * @throws IOException
	 */
	private byte[] toPictureBytes(BufferedImage bufferImg, String fileSuffix) throws IOException {
		try (ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream()) {
			boolean success = ImageIO.write(bufferImg, fileSuffix, byteArrayOut);
			if (!success) {
				throw new IOException("不支持的图片格式：" + fileSuffix);
			}
			return byteArrayOut.toByteArray();
		}
	}

	/**
	 * 获取POI图片类型。
	 * @param fileSuffix
	 * @return
	 */
	private int getPoiImageType(String fileSuffix) {
		return "png".equals(fileSuffix) ? Workbook.PICTURE_TYPE_PNG : Workbook.PICTURE_TYPE_JPEG;
	}

	/**
	 * 从路径或URL中解析图片后缀，自动去掉查询参数。
	 * @param path
	 * @return
	 */
	private String getImageSuffix(String path) {
		if (path==null || path.length()==0) {
			return "png";
		}
		int queryIndex = path.indexOf("?");
		if (queryIndex>=0) {
			path = path.substring(0, queryIndex);
		}
		int hashIndex = path.indexOf("#");
		if (hashIndex>=0) {
			path = path.substring(0, hashIndex);
		}
		int dotIndex = path.lastIndexOf(".");
		if (dotIndex<0 || dotIndex==path.length()-1) {
			return "png";
		}

		return path.substring(dotIndex + 1).toLowerCase();
	}

	/**
	 * 图片实际展示尺寸。
	 */
	private static class ImageSize {
		private final int width;
		private final int height;

		private ImageSize(int width, int height) {
			this.width = width;
			this.height = height;
		}

		private int getWidth() {
			return width;
		}

		private int getHeight() {
			return height;
		}
	}

	/**
	 * 图片字节和格式。
	 */
	private static class ImageData {
		private final byte[] bytes;
		private final String fileSuffix;

		private ImageData(byte[] bytes, String fileSuffix) {
			this.bytes = bytes;
			this.fileSuffix = fileSuffix;
		}

		private byte[] getBytes() {
			return bytes;
		}

		private String getFileSuffix() {
			return fileSuffix;
		}
	}

	/**
	 * 替换模板中的占位符（占位符独占一格单元格）
	 * @param cell
	 * @param obj
	 */
	public void setValue(Cell cell, Object obj) {
		if (obj==null) {
			cell.setBlank();
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

			Workbook wb = cell.getSheet().getWorkbook();
			CellStyle base = cell.getCellStyle();
			CellStyle cached = getOrCreateStyleWithFont(wb, base, font);

			cell.setCellValue(font.getText());
			cell.setCellStyle(cached);
		}
		// 自定义图片
		else if (obj instanceof Picture) {
			cell.setCellValue("");
			Picture picture = (Picture) obj;
			this.setImage(cell, picture);
		}
		// 多张图片
		else if (obj instanceof List && this.isImageList((List<?>) obj)) {
			cell.setCellValue("");
			this.setImages(cell, (List<?>) obj);
		}
		else {
			cell.setCellValue(obj.toString());
		}
	}

	/**
	 * 判断集合是否可以按图片集合处理。
	 * @param list
	 * @return
	 */
	private boolean isImageList(List<?> list) {
		if (list==null || list.isEmpty()) {
			return false;
		}
		boolean hasPicture = false;
		for (Object obj : list) {
			if (!(obj instanceof Picture) && !(obj instanceof String) && !(obj instanceof byte[])) {
				return false;
			}
			if (obj instanceof Picture || obj instanceof byte[]) {
				hasPicture = true;
			}
		}

		return hasPicture;
	}

	/**
	 * 自定义字体缓存key
	 * @param f
	 * @return
	 */
	private String fontKey(Font f) {
	    return (f.getColor() == null ? "" : f.getColor().trim().toUpperCase()) + "|" +
	           (f.getFontFamily() == null ? "" : f.getFontFamily().trim()) + "|" +
	           (f.getFontSize() == null ? "" : f.getFontSize()) + "|" +
	           f.getBold() + "|" + f.getItalic() + "|" + f.getStrike();
	}

	/**
	 * 得到Poi字体
	 * @param wb
	 * @param f
	 * @return
	 */
	private org.apache.poi.ss.usermodel.Font getOrCreatePoiFont(Workbook wb, Font f) {
	    String key = fontKey(f);
	    return poiFontCache.computeIfAbsent(key, k -> {
	        XSSFFont pf = (XSSFFont) wb.createFont();
	        if (f.getFontFamily() != null) pf.setFontName(f.getFontFamily());
	        if (f.getFontSize() != null) pf.setFontHeightInPoints(f.getFontSize().shortValue());
	        pf.setBold(f.getBold());
	        pf.setItalic(f.getItalic());
	        pf.setStrikeout(f.getStrike());

	        if (f.getColor() != null && f.getColor().length() > 0) {
	            byte[] rgb = hexToRgb(f.getColor());
	            pf.setColor(new XSSFColor(rgb, null));
	        }
	        return pf;
	    });
	}

	/**
	 * 字体样式
	 * @param wb
	 * @param baseStyle
	 * @param f
	 * @return
	 */
	private CellStyle getOrCreateStyleWithFont(Workbook wb, CellStyle baseStyle, Font f) {
	    String key = System.identityHashCode(baseStyle) + "||" + fontKey(f);
	    return styleCache.computeIfAbsent(key, k -> {
	        CellStyle ns = wb.createCellStyle();
	        ns.cloneStyleFrom(baseStyle);
	        ns.setFont(getOrCreatePoiFont(wb, f));
	        return ns;
	    });
	}

	/**
	 * 颜色转换
	 * @param hex
	 * @return
	 */
	private static byte[] hexToRgb(String hex) {
	    String s = hex.trim();
	    if (s.startsWith("#")) s = s.substring(1);
	    if (s.length() != 6) throw new IllegalArgumentException("color must be RRGGBB, but: " + hex);
	    return new byte[] {
	        (byte) Integer.parseInt(s.substring(0, 2), 16),
	        (byte) Integer.parseInt(s.substring(2, 4), 16),
	        (byte) Integer.parseInt(s.substring(4, 6), 16)
	    };
	}

	/**
	 * 规范化图片后缀，避免未填写或填写jpeg时 ImageIO / POI 类型判断不一致。
	 * @param fileSuffix
	 * @return
	 */
	private String normalizeImageSuffix(String fileSuffix) {
		if (fileSuffix==null || fileSuffix.length()==0) {
			return "png";
		}
		fileSuffix = fileSuffix.toLowerCase();
		if ("jpeg".equals(fileSuffix)) {
			return "jpg";
		}
		if (!"jpg".equals(fileSuffix) && !"png".equals(fileSuffix)) {
			return "png";
		}

		return fileSuffix;
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
			Object obj = PropertyHandler.getValue(param, placeholder);
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
			Workbook wb = cell.getSheet().getWorkbook();

			for (Map.Entry<String, Font> entry : map.entrySet()) {
				String key = entry.getKey();
		        if (key == null || key.isEmpty()) {
		        	continue;
		        }

		        XSSFFont poiFont = (XSSFFont) getOrCreatePoiFont(wb, entry.getValue());

		        int from = 0;
		        while (from < cellValue.length()) {
		            int start = cellValue.indexOf(key, from);
		            if (start < 0) {
		                break;
		            }
		            int end = start + key.length();
		            richString.applyFont(start, end, poiFont);

		            // 推进搜索起点，确保不会重复命中同一位置
		            from = end;
		        }
			}

			cell.setCellValue(richString);
		}
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
			throw new RuntimeException("Failed to set cell style", e);
		}
	}

}
