package cn.javaex.officejj.excel.help;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import cn.javaex.officejj.common.entity.Picture;
import cn.javaex.officejj.excel.ExcelUtils;
import cn.javaex.officejj.excel.annotation.ExcelCell;
import cn.javaex.officejj.excel.annotation.FormatValidation;
import cn.javaex.officejj.excel.annotation.NotEmpty;
import cn.javaex.officejj.excel.exception.ExcelValidationException;
import cn.javaex.officejj.excel.function.DefaultExcelValueConverter;
import cn.javaex.officejj.excel.function.ExcelValueConverter;

import javax.imageio.ImageIO;

/**
 * 读取Excel
 *
 * @author 陈霓清
 */
public class SheetReadHelper extends SheetHelper {

	/**
	 * 读取sheet
	 * @param <T>
	 * @param sheet
	 * @param clazz     自定义实体类
	 * @param rowNum    从第几行开始读取（从0开始计算）
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
	@Override
	public <T> List<T> read(Sheet sheet, Class<T> clazz, int rowNum) throws Exception {
		List<T> list = new ArrayList<T>();
		StringBuffer sb = new StringBuffer();
		Map<Integer, List<String>> rowErrorMap = new LinkedHashMap<Integer, List<String>>();

		Field[] fieldArr = clazz.getDeclaredFields();

		// 1.0 解析注解
		this.readAnnotation(fieldArr);
		// 合并单元格读取器。导入明细表时，省/市/班级等父级列经常做纵向合并，
		// POI 只有合并区域左上角单元格有值，其他格子需要回读左上角内容。
		MergedCellValueReader mergedCellValueReader = new MergedCellValueReader(sheet);
		// 图片读取器。WPS/Excel 保存的 xlsx 图片会落在 Drawing 关系里，按锚点左上角单元格归档。
		ImageCellReader imageCellReader = new ImageCellReader(sheet);

		// 2.0 遍历数据
		for (Row row : sheet) {
			// 跳过表头
			if (row.getRowNum()<rowNum) {
				continue;
			}
			// 模板中预设下拉框、样式或保护区域时，POI 可能会把后续空行也作为物理行返回。
			// 这些行没有用户填写的数据，应直接跳过，避免触发必填项校验。
			if (this.isBlankDataRow(row, fieldArr, imageCellReader)) {
				continue;
			}

			// 遍历每一列
			T entity = null;
			int len = fieldArr.length;
			for (int i=0; i<len; i++) {
				// 根据对象类型设置值
				Field field = fieldArr[i];
				field.setAccessible(true);    // 设置类的私有属性可访问

				ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
				int sort = this.getFieldSort(field, i);

				if (this.isImageField(excelCell)) {
					List<Picture> pictureList = imageCellReader.getPictures(row.getRowNum(), sort);
					if (pictureList.isEmpty()) {
						this.validateEmptyCell(field, sort, row, sb, rowErrorMap);
						continue;
					}

					// 如果实例不存在则新建
					if (entity==null) {
						entity = clazz.getDeclaredConstructor().newInstance();
					}

					try {
						field.set(entity, this.convertImageValue(field, pictureList));
					} catch (Exception e) {
						FormatValidation formatValidation = field.getAnnotation(FormatValidation.class);
						int excelRowNum = row.getRowNum() + 1;
						String message = "第" + (sort+1) + "列，" + (formatValidation==null ? "图片字段类型不正确：" + e.getMessage() : formatValidation.value());
						sb.append("第" + excelRowNum + "行，" + message);
						sb.append("<br/>");
						this.addRowError(rowErrorMap, excelRowNum, message);
					}
					continue;
				}

				// 获取该列的值
				String cellValue = mergedCellValueReader.getCellValue(row, sort);
				// 默认值
				if (cellValue.length()==0) {
					if (excelCell!=null && excelCell.defaultValue().length()>0) {
						cellValue = excelCell.defaultValue();
					}
				}
				// 必填项校验
				if (cellValue.length()==0) {
					this.validateEmptyCell(field, sort, row, sb, rowErrorMap);
					continue;
				}
				// 值替换
				if (excelCell!=null && excelCell.replace().length>0) {
					Map<String, String> map = (Map<String, String>) replaceMap.get(String.valueOf(sort));
					if (map!=null && map.get(cellValue)!=null) {
						cellValue = map.get(cellValue.toString());
					}
				}

				// 如果实例不存在则新建
				if (entity==null) {
					entity = clazz.getDeclaredConstructor().newInstance();
				}

				try {
					field.set(entity, this.convertCellValue(field, excelCell, cellValue, sort));
				} catch (Exception e) {
					FormatValidation formatValidation = field.getAnnotation(FormatValidation.class);
					int excelRowNum = row.getRowNum() + 1;
					String message = "第" + (sort+1) + "列，" + (formatValidation==null ? "数据格式不正确：" + e.getMessage() : formatValidation.value());
					sb.append("第" + excelRowNum + "行，" + message);
					sb.append("<br/>");
					this.addRowError(rowErrorMap, excelRowNum, message);
					continue;
				}
			}

			// 把每一行的实体对象加入list
			if (entity!=null) {
				list.add(entity);
			}
		}

		if (!"".equals(sb.toString())) {
			throw new ExcelValidationException(sb.toString(), rowErrorMap);
		}

		return list;
	}

	/**
	 * 判断当前行是否为空白数据行。
	 * 模板行可能只有下拉框、样式、边框等元数据，没有真实填写内容；这类行不应参与导入校验。
	 * 判断时不能使用 ExcelCell.defaultValue，否则完全空白的模板行会被默认值误认为有效数据。
	 * @param row 当前行
	 * @param fieldArr 导入实体字段
	 * @param imageCellReader 图片读取器
	 * @return 是否为空白数据行
	 */
	private boolean isBlankDataRow(Row row, Field[] fieldArr, ImageCellReader imageCellReader) {
		int len = fieldArr.length;
		for (int i=0; i<len; i++) {
			Field field = fieldArr[i];
			int sort = this.getFieldSort(field, i);
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (this.isImageField(excelCell) && imageCellReader.hasPictures(row.getRowNum(), sort)) {
				return false;
			}

			String cellValue = ExcelUtils.getCellValue(row.getCell(sort));
			if (cellValue!=null && cellValue.trim().length()>0) {
				return false;
			}
		}

		return true;
	}

	/**
	 * xlsx 图片读取器。
	 * 仅完整 Workbook 读取流程能拿到 Drawing 关系；SAX 低内存读取仍按文本数据处理。
	 */
	private static class ImageCellReader {
		private final Map<String, List<Picture>> pictureMap = new HashMap<String, List<Picture>>();

		private ImageCellReader(Sheet sheet) {
			this.readXssfPictures(sheet);
			this.readHssfPictures(sheet);
		}

		/**
		 * 读取 XSSF/WPS xlsx 中的图片，按锚点左上角单元格归档。
		 * @param sheet 当前Sheet
		 */
		private void readXssfPictures(Sheet sheet) {
			if (!(sheet instanceof XSSFSheet)) {
				return;
			}

			XSSFDrawing drawing = ((XSSFSheet) sheet).getDrawingPatriarch();
			if (drawing==null) {
				return;
			}

			for (XSSFShape shape : drawing.getShapes()) {
				if (!(shape instanceof XSSFPicture)) {
					continue;
				}

				XSSFPicture picture = (XSSFPicture) shape;
				XSSFAnchor anchor = picture.getAnchor();
				if (!(anchor instanceof XSSFClientAnchor)) {
					continue;
				}

				XSSFClientAnchor clientAnchor = (XSSFClientAnchor) anchor;
				Picture importPicture = this.toPicture(picture);
				if (importPicture!=null) {
					this.addPicture(clientAnchor.getRow1(), clientAnchor.getCol1(), importPicture);
				}
			}
		}

		/**
		 * 把 POI 图片数据转换成 officejj 的 Picture 对象，保留字节、后缀和原图尺寸。
		 * @param xssfPicture POI 图片对象
		 * @return 图片对象
		 */
		private Picture toPicture(XSSFPicture xssfPicture) {
			XSSFPictureData pictureData = xssfPicture.getPictureData();
			if (pictureData==null || pictureData.getData()==null || pictureData.getData().length==0) {
				return null;
			}

			return this.toPicture(pictureData.getData(), pictureData.suggestFileExtension());
		}

		/**
		 * 读取 HSSF/WPS xls 中的图片，兼容客户仍使用旧版 Excel 格式的场景。
		 * @param sheet 当前Sheet
		 */
		private void readHssfPictures(Sheet sheet) {
			if (!(sheet instanceof HSSFSheet)) {
				return;
			}

			HSSFPatriarch drawing = ((HSSFSheet) sheet).getDrawingPatriarch();
			if (drawing==null) {
				return;
			}

			for (HSSFShape shape : drawing.getChildren()) {
				if (!(shape instanceof HSSFPicture)) {
					continue;
				}

				HSSFPicture picture = (HSSFPicture) shape;
				if (!(picture.getAnchor() instanceof HSSFClientAnchor)) {
					continue;
				}

				HSSFPictureData pictureData = picture.getPictureData();
				if (pictureData==null || pictureData.getData()==null || pictureData.getData().length==0) {
					continue;
				}

				HSSFClientAnchor anchor = (HSSFClientAnchor) picture.getAnchor();
				Picture importPicture = this.toPicture(pictureData.getData(), pictureData.suggestFileExtension());
				if (importPicture!=null) {
					this.addPicture(anchor.getRow1(), anchor.getCol1(), importPicture);
				}
			}
		}

		/**
		 * 把图片字节转换成 officejj 的 Picture 对象。
		 * @param data 图片字节
		 * @param fileSuffix 图片后缀
		 * @return 图片对象
		 */
		private Picture toPicture(byte[] data, String fileSuffix) {
			if (data==null || data.length==0) {
				return null;
			}

			Picture picture = new Picture();
			picture.setData(data);
			picture.setFileSuffix(fileSuffix);
			picture.setDescription(fileSuffix==null || fileSuffix.length()==0 ? "import-image" : "import-image." + fileSuffix);
			try {
				BufferedImage image = ImageIO.read(new ByteArrayInputStream(data));
				if (image!=null) {
					picture.setWidth(image.getWidth());
					picture.setHeight(image.getHeight());
				}
			} catch (Exception ignore) {
				// 部分 WPS 文件可能包含 emf/wmf 等 ImageIO 不能解析的格式，字节内容仍然保留给业务层处理。
			}
			return picture;
		}

		private void addPicture(int rowIndex, int colIndex, Picture picture) {
			String key = this.key(rowIndex, colIndex);
			List<Picture> list = this.pictureMap.get(key);
			if (list==null) {
				list = new ArrayList<Picture>();
				this.pictureMap.put(key, list);
			}
			list.add(picture);
		}

		private List<Picture> getPictures(int rowIndex, int colIndex) {
			List<Picture> list = this.pictureMap.get(this.key(rowIndex, colIndex));
			if (list==null || list.isEmpty()) {
				return new ArrayList<Picture>();
			}
			return new ArrayList<Picture>(list);
		}

		private boolean hasPictures(int rowIndex, int colIndex) {
			List<Picture> list = this.pictureMap.get(this.key(rowIndex, colIndex));
			return list!=null && !list.isEmpty();
		}

		private String key(int rowIndex, int colIndex) {
			return rowIndex + "_" + colIndex;
		}
	}

	/**
	 * 合并单元格读取器。
	 * 普通导入时把合并区域内的空白单元格，按左上角单元格的值进行回填。
	 * 空白行判断仍然只看当前行真实填写的单元格，避免整块合并区域的空白下半部分被误导入。
	 */
	private static class MergedCellValueReader {
		private final Sheet sheet;
		private final Map<Integer, List<CellRangeAddress>> rowMergedRegionMap = new HashMap<Integer, List<CellRangeAddress>>();
		private final Map<String, String> valueCache = new HashMap<String, String>();

		private MergedCellValueReader(Sheet sheet) {
			this.sheet = sheet;
			this.initRowMergedRegionMap(sheet);
		}

		/**
		 * 按行缓存合并区域，避免导入每个空单元格时都扫描整张表的所有合并区域。
		 * @param sheet 当前Sheet
		 */
		private void initRowMergedRegionMap(Sheet sheet) {
			List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
			for (CellRangeAddress region : mergedRegions) {
				for (int rowIndex=region.getFirstRow(); rowIndex<=region.getLastRow(); rowIndex++) {
					List<CellRangeAddress> regionList = this.rowMergedRegionMap.get(rowIndex);
					if (regionList==null) {
						regionList = new ArrayList<CellRangeAddress>();
						this.rowMergedRegionMap.put(rowIndex, regionList);
					}
					regionList.add(region);
				}
			}
		}

		/**
		 * 读取单元格文本。当前单元格为空且属于合并区域时，返回合并区域左上角的文本。
		 * @param row 当前行
		 * @param colIndex 列索引，从0开始
		 * @return 单元格文本
		 */
		private String getCellValue(Row row, int colIndex) {
			if (row==null) {
				return "";
			}

			String cellValue = ExcelUtils.getCellValue(row.getCell(colIndex));
			if (cellValue!=null && cellValue.length()>0) {
				return cellValue;
			}

			int rowIndex = row.getRowNum();
			String key = rowIndex + "_" + colIndex;
			if (this.valueCache.containsKey(key)) {
				return this.valueCache.get(key);
			}

			CellRangeAddress mergedRegion = this.getMergedRegion(rowIndex, colIndex);
			if (mergedRegion==null) {
				this.valueCache.put(key, "");
				return "";
			}

			Row firstRow = this.sheet.getRow(mergedRegion.getFirstRow());
			String mergedValue = firstRow==null ? "" : ExcelUtils.getCellValue(firstRow.getCell(mergedRegion.getFirstColumn()));
			this.valueCache.put(key, mergedValue);
			return mergedValue;
		}

		/**
		 * 查找指定单元格所在的合并区域。
		 * @param rowIndex 行索引，从0开始
		 * @param colIndex 列索引，从0开始
		 * @return 合并区域，未命中时返回null
		 */
		private CellRangeAddress getMergedRegion(int rowIndex, int colIndex) {
			List<CellRangeAddress> regionList = this.rowMergedRegionMap.get(rowIndex);
			if (regionList==null || regionList.isEmpty()) {
				return null;
			}

			for (CellRangeAddress region : regionList) {
				if (region.isInRange(rowIndex, colIndex)) {
					return region;
				}
			}

			return null;
		}
	}

	/**
	 * 获取字段对应的列索引。
	 * 未配置 ExcelCell.sort 时，沿用字段声明顺序，保持旧版本导入行为。
	 * @param field 字段
	 * @param defaultIndex 默认列索引
	 * @return 从0开始的列索引
	 */
	private int getFieldSort(Field field, int defaultIndex) {
		ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
		return excelCell==null || excelCell.sort()==0 ? defaultIndex : (excelCell.sort() - 1);
	}

	/**
	 * 记录某一行的导入错误。
	 * @param rowErrorMap 行错误集合
	 * @param excelRowNum Excel行号，从1开始
	 * @param message 错误信息
	 */
	private void addRowError(Map<Integer, List<String>> rowErrorMap, int excelRowNum, String message) {
		List<String> list = rowErrorMap.get(excelRowNum);
		if (list==null) {
			list = new ArrayList<String>();
			rowErrorMap.put(excelRowNum, list);
		}
		list.add(message);
	}

	/**
	 * 必填项校验。
	 * 文本和图片字段共用同一套错误收集逻辑，便于前端统一展示第几行、第几列错误。
	 * @param field 字段
	 * @param sort 列索引，从0开始
	 * @param row 当前行
	 * @param sb 错误文本
	 * @param rowErrorMap 行错误集合
	 */
	private void validateEmptyCell(Field field, int sort, Row row, StringBuffer sb, Map<Integer, List<String>> rowErrorMap) {
		NotEmpty notEmpty = field.getAnnotation(NotEmpty.class);
		if (notEmpty==null) {
			return;
		}

		int excelRowNum = row.getRowNum() + 1;
		String message = "第" + (sort+1) + "列，" + notEmpty.value();
		sb.append("第" + excelRowNum + "行，" + message);
		sb.append("<br/>");
		this.addRowError(rowErrorMap, excelRowNum, message);
	}

	/**
	 * 将单元格文本转换成字段类型。
	 * 优先执行用户在 ExcelCell.converter 中配置的转换器；未配置时使用内置基础类型转换。
	 * @param field 字段
	 * @param excelCell 字段注解
	 * @param cellValue 单元格文本
	 * @param sort 列索引，从0开始
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings({"unchecked", "rawtypes"})
	private Object convertCellValue(Field field, ExcelCell excelCell, String cellValue, int sort) throws Exception {
		if (this.hasCustomConverter(excelCell)) {
			// 用户显式配置 converter 时，必须先交给业务转换器处理。
			// 例如 Excel 单元格为“男”，字段类型为 Integer，不能先走内置数字解析。
			ExcelValueConverter converter = excelCell.converter().getDeclaredConstructor().newInstance();
			return converter.convert(cellValue, field, excelCell);
		}

		Class<?> fieldType = field.getType();
		if (fieldType==String.class) {
			return cellValue;
		}
		if (fieldType==Integer.class || fieldType==Integer.TYPE) {
			return Double.valueOf(String.valueOf(cellValue)).intValue();
		}
		if (fieldType==Long.class || fieldType==Long.TYPE) {
			return new BigDecimal(cellValue).longValue();
		}
		if (fieldType==Double.class || fieldType==Double.TYPE) {
			return Double.valueOf(cellValue);
		}
		if (fieldType==Float.class || fieldType==Float.TYPE) {
			return Float.valueOf(cellValue);
		}
		if (fieldType==BigDecimal.class) {
			return new BigDecimal(cellValue);
		}
		if (fieldType==Boolean.class || fieldType==Boolean.TYPE) {
			return Boolean.valueOf(cellValue);
		}
		if (fieldType.isEnum()) {
			for (Object enumValue : fieldType.getEnumConstants()) {
				Enum item = (Enum) enumValue;
				if (item.name().equals(cellValue) || item.toString().equals(cellValue)) {
					return item;
				}
			}
			throw new IllegalArgumentException("无法匹配枚举值：" + cellValue);
		}
		if (fieldType==LocalDateTime.class) {
			SimpleDateFormat sdf = (SimpleDateFormat) formatMap.get(String.valueOf(sort));
			if (sdf!=null) {
				cellValue = this.timestampToDateString(sdf, cellValue);
				Instant instant = sdf.parse(cellValue).toInstant();
				return LocalDateTime.ofInstant(instant, ZoneId.systemDefault());
			}
			if (cellValue.length()==13) {
				Instant instant = Instant.ofEpochMilli(Long.parseLong(cellValue));
				return LocalDateTime.ofInstant(instant, ZoneId.systemDefault());
			}
			return LocalDateTime.parse(cellValue);
		}
		if (fieldType==LocalDate.class) {
			DateTimeFormatter dtf = (DateTimeFormatter) formatMap.get(String.valueOf(sort));
			if (dtf!=null) {
				if (cellValue.length()==13) {
					LocalDateTime ofInstant = LocalDateTime.ofInstant(Instant.ofEpochSecond(Long.parseLong(cellValue) / 1000L), TimeZone.getDefault().toZoneId());
					return ofInstant.toLocalDate();
				}
				return LocalDate.parse(cellValue, dtf);
			}
			if (cellValue.length()==13) {
				LocalDateTime ofInstant = LocalDateTime.ofInstant(Instant.ofEpochMilli(Long.parseLong(cellValue)), TimeZone.getDefault().toZoneId());
				return ofInstant.toLocalDate();
			}
			return LocalDate.parse(cellValue);
		}
		if (fieldType==Date.class) {
			SimpleDateFormat sdf = (SimpleDateFormat) formatMap.get(String.valueOf(sort));
			if (sdf!=null) {
				cellValue = this.timestampToDateString(sdf, cellValue);
				return sdf.parse(cellValue);
			}
			if (cellValue.length()==13) {
				return new Date(Long.parseLong(cellValue));
			}
			throw new IllegalArgumentException("日期字段未配置format");
		}
		return null;
	}

	/**
	 * 将图片列表转换成导入 DTO 字段类型。
	 * 单图片字段建议使用 Picture；多图片字段建议使用 List<Picture>。
	 * @param field 字段
	 * @param pictureList 当前单元格锚定的图片集合
	 * @return 字段值
	 */
	private Object convertImageValue(Field field, List<Picture> pictureList) {
		Class<?> fieldType = field.getType();
		if (fieldType==Picture.class) {
			return pictureList.get(0);
		}
		if (List.class.isAssignableFrom(fieldType)) {
			return new ArrayList<Picture>(pictureList);
		}
		if (fieldType==byte[].class) {
			return pictureList.get(0).getData();
		}
		if (fieldType==byte[][].class) {
			byte[][] bytesArr = new byte[pictureList.size()][];
			for (int i=0; i<pictureList.size(); i++) {
				bytesArr[i] = pictureList.get(i).getData();
			}
			return bytesArr;
		}

		throw new IllegalArgumentException("图片字段仅支持 Picture、List<Picture>、byte[]、byte[][]");
	}

	/**
	 * 判断字段是否为图片导入字段。
	 * @param excelCell 字段注解
	 * @return 是否图片字段
	 */
	private boolean isImageField(ExcelCell excelCell) {
		return excelCell!=null && excelCell.type()!=null && excelCell.type().contains("image");
	}

	/**
	 * 判断字段是否配置了自定义导入转换器。
	 * 默认转换器只用于占位，表示继续使用 officejj 内置的基础类型转换逻辑。
	 * @param excelCell 字段注解
	 * @return 是否需要优先执行用户自定义转换器
	 */
	private boolean hasCustomConverter(ExcelCell excelCell) {
		return excelCell!=null && excelCell.converter()!=DefaultExcelValueConverter.class;
	}

	/**
	 * 时间戳转日期格式
	 * @param sdf
	 * @param cellValue
	 * @return
	 */
	private String timestampToDateString(SimpleDateFormat sdf, String cellValue) {
		if (cellValue.length()==13) {	// 时间戳
			return sdf.format(Long.parseLong(cellValue));
		}

		return cellValue;
	}

	/**
	 * 解析注解
	 * @param <T>
	 * @param fieldArr
	 */
	private void readAnnotation(Field[] fieldArr) {
		replaceMap.clear();
		formatMap.clear();

		for (int i=0; i<fieldArr.length; i++) {
			Field field = fieldArr[i];
			// 设置类的私有属性可访问
			field.setAccessible(true);

			// 得到每一个成员变量上的 ExcelCell 注解
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (excelCell==null) {
				continue;
			}

			int sort = excelCell.sort()==0 ? i : (excelCell.sort() - 1);

			// 设置值替换属性
			String[] replaceArr = excelCell.replace();
			if (replaceArr.length>0) {
				Map<String, String> map = new HashMap<String, String>();
				// {"男_1", "女_0"}
				for (String replace : replaceArr) {
					// 男_1
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
				if (field.getType()==LocalDateTime.class || field.getType()==Date.class) {
					SimpleDateFormat sdf = new SimpleDateFormat(format);
					formatMap.put(String.valueOf(sort), sdf);
				}
				else if (field.getType()==LocalDate.class) {
					DateTimeFormatter dtf = DateTimeFormatter.ofPattern(format);
					formatMap.put(String.valueOf(sort), dtf);
				}
			}
		}
	}

}
