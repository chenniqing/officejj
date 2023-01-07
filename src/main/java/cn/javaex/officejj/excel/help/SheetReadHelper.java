package cn.javaex.officejj.excel.help;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import cn.javaex.officejj.excel.ExcelUtils;
import cn.javaex.officejj.excel.annotation.ExcelCell;
import cn.javaex.officejj.excel.annotation.FormatValidation;
import cn.javaex.officejj.excel.annotation.NotEmpty;
import cn.javaex.officejj.excel.exception.ExcelValidationException;

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
		
		Field[] fieldArr = clazz.getDeclaredFields();
		
		// 1.0 解析注解
		this.readAnnotation(fieldArr);
		
		// 2.0 遍历数据
		for (Row row : sheet) {
			// 跳过表头
			if (row.getRowNum()<rowNum) {
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
				
				// 获取该列的值
				String cellValue = ExcelUtils.getCellValue(row.getCell(i));
				// 默认值
				if (cellValue.length()==0) {
					if (excelCell!=null && excelCell.defaultValue().length()>0) {
						cellValue = excelCell.defaultValue();
					}
				}
				// 必填项校验
				if (cellValue.length()==0) {
					NotEmpty notEmpty = field.getAnnotation(NotEmpty.class);
					if (notEmpty!=null) {
						sb.append("第" + (row.getRowNum()+1) + "行，第" + (i+1) + "列，" + notEmpty.value());
						sb.append("<br/>"); 
					}
					
					continue;
				}
				// 值替换
				if (excelCell!=null && excelCell.replace().length>0) {
					Map<String, String> map = (Map<String, String>) replaceMap.get(String.valueOf(i));
					if (map.get((String) cellValue)!=null) {
						cellValue = map.get(cellValue.toString());
					}
				}
				
				// 如果实例不存在则新建
				if (entity==null) {
					entity = clazz.newInstance();
				}
				
				Class<?> fieldType = field.getType();
				if (fieldType==String.class) {
					field.set(entity, cellValue);
				}
				else if (fieldType==Integer.class || fieldType==Integer.TYPE) {
					field.set(entity, Double.valueOf(String.valueOf(cellValue)).intValue());
				}
				else if (fieldType==Long.class || fieldType==Long.TYPE) {
					field.set(entity, Long.valueOf(cellValue));
				}
				else if (fieldType==Double.class || fieldType==Double.TYPE) {
					field.set(entity, Double.valueOf(cellValue));
				}
				else if (fieldType==Float.class || fieldType==Float.TYPE) {
					field.set(entity, Float.valueOf(cellValue));
				}
				else if (fieldType==LocalDateTime.class || fieldType==LocalDate.class || fieldType==Date.class) {
					try {
						if (fieldType==LocalDateTime.class) {
							SimpleDateFormat sdf = (SimpleDateFormat) formatMap.get(String.valueOf(i));
							if (sdf!=null) {
								cellValue = this.timestampToDateString(sdf, cellValue);
								
								Instant instant = sdf.parse(cellValue).toInstant();
								ZoneId zone = ZoneId.systemDefault();
								LocalDateTime localDateTime = LocalDateTime.ofInstant(instant, zone);
								
								field.set(entity, localDateTime);
							}
						}
						else if (fieldType==LocalDate.class) {
							DateTimeFormatter dtf = (DateTimeFormatter) formatMap.get(String.valueOf(i));
							if (dtf!=null) {
								LocalDate ld = null;
								if (cellValue.length()==13) {	// 时间戳
									LocalDateTime ofInstant = LocalDateTime.ofInstant(Instant.ofEpochSecond(Long.parseLong(cellValue) / 1000L), TimeZone.getDefault().toZoneId());
									ld = ofInstant.toLocalDate();
								} else {
									ld = LocalDate.parse(cellValue, dtf);
								}
								
								field.set(entity, ld);
							}
						}
						else if (fieldType==Date.class) {
							SimpleDateFormat sdf = (SimpleDateFormat) formatMap.get(String.valueOf(i));
							if (sdf!=null) {
								cellValue = this.timestampToDateString(sdf, cellValue);
								
								field.set(entity, sdf.parse(cellValue));
							}
						}
					} catch (Exception e) {
						// 格式化校验
						FormatValidation formatValidation = field.getAnnotation(FormatValidation.class);
						if (formatValidation!=null) {
							sb.append("第" + (row.getRowNum()+1) + "行，第" + (i+1) + "列，" + formatValidation.value());
							sb.append("<br/>"); 
							
							continue;
						} else {
							field.set(entity, null);
						}
					}
				}
				else {
					field.set(entity, null);
				}
			}
			
			// 把每一行的实体对象加入list
			if (entity!=null) {
				list.add(entity);
			}
		}
		
		if (!"".equals(sb.toString())) {
			throw new ExcelValidationException(sb.toString());
		}
		
		return list;
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
					String[] arr = replace.split("_");
					map.put(arr[0], arr[1]);
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
