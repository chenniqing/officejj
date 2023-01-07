package cn.javaex.officejj.excel.help;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import cn.javaex.officejj.common.util.MapHandler;
import cn.javaex.officejj.excel.ExcelUtils;

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
			
			int len = list.size();
			// 2.0 插入并复制行
			List<Map<String, Integer[]>> tempPlaceholders = entry.getValue();
			Integer[] tempResult = MapHandler.getFirstOrNull(tempPlaceholders.get(0));
			super.insertRow(sheet, tempResult[0], len-1);
			
			// 3.0 遍历取出每一条数据并设置值
			for (int i=0; i<len; i++) {
				// 获得传入的每一条数据
				Map<String, Object> dataMap = list.get(i);
				
				// 获得一整行的占位符记录
				List<Map<String, Integer[]>> placeholders = entry.getValue();
				for (Map<String, Integer[]> placeholder : placeholders) {
					// 获取该行的每一个单元格记录
					for (Map.Entry<String, Integer[]> placeholderMap : placeholder.entrySet()) {
						String attributeKey = placeholderMap.getKey();         // 属性Key
						Integer[] coordinate = placeholderMap.getValue();      // Cell坐标（行索引、列索引）
						
						// 设置值替换
						Cell cell = sheet.getRow(coordinate[0]+i).getCell(coordinate[1]);
						cellHelper.setValue(cell, dataMap.get(attributeKey));
					}
				}
			}
		}
	}
	
}
