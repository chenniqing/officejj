package cn.javaex.officejj.word.help;

import java.io.ByteArrayInputStream;
import java.io.ObjectInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc.Enum;

import cn.javaex.officejj.common.entity.Font;
import cn.javaex.officejj.common.entity.Table;

/**
 * 表格
 * 
 * @author 陈霓清
 */
public class TableHelper extends ParagraphHelper {

	/**
	 * 替换表格变量
	 * @param doc
	 * @param param
	 */
	public void replaceTable(XWPFDocument word, Map<String, Object> param) {
		Iterator<XWPFTable> iterator = word.getTablesIterator();
		XWPFTable table = null;
		
		while (iterator.hasNext()) {
			table = iterator.next();
			if (table.getRows().size()>0) {
				if (matcher(table.getText()).find()) {
					boolean flag = true;
					String key = "";
					int index = 1;
					
					jump:
					for (int i=0; i<table.getRows().size(); i++) {
						XWPFTableRow row = table.getRows().get(i);
						for (XWPFTableCell cell : row.getTableCells()) {
							for (XWPFParagraph paragraph : cell.getParagraphs()) {
								try {
									super.replaceParagraph(paragraph, param);
								} catch (Exception e) {
									// 此处表示是自动插入循环表格数据
									flag = false;
									key = e.getMessage();
									
									// 设置指定变量为指定值
									super.replaceParagraph(paragraph, "");
									
									// 第几行开始是数据（从0开始计）
									index = i;
									
									break jump;
								}
							}
						}
					}
					
					if (!flag) {
						// 为表格插入数据
						if (param.get(key)!=null && (param.get(key) instanceof Table)) {
							Table tableSetting = (Table) param.get(key);
							
							// 1.0 插入数据
							List<String[]> dataList = tableSetting.getDataList();
							if (dataList!=null && dataList.isEmpty()==false) {
								this.insertTable(table, dataList, index);
							}
							
							// 2.0 单元格合并列
							List<int[]> mergeColList = tableSetting.getMergeColList();
							if (mergeColList!=null && mergeColList.isEmpty()==false) {
								for (int[] mergeColArr : mergeColList) {
									this.mergeCol(table, mergeColArr[0]-1, mergeColArr[1]-1, mergeColArr[2]-1);
								}
							}
							
							// 3.0 单元格合并行
							List<int[]> mergeRowList = tableSetting.getMergeRowList();
							if (mergeRowList!=null && mergeRowList.isEmpty()==false) {
								for (int[] mergeRowArr : mergeRowList) {
									this.mergeRow(table, mergeRowArr[0]-1, mergeRowArr[1]-1, mergeRowArr[2]-1);
								}
							}
						}
					}
				}
			}
		}
	}
	
	/**
	 * 插入动态表格数据
	 * @param table 需要插入数据的表格
	 * @param dataList 插入数据集合
	 * @param index 表头行数/第一行数据行所在的索引位置
	 */
	private void insertTable(XWPFTable table, List<String[]> dataList, int index) {
		RunHelper runHelper = new RunHelper();
		
		// 创建行，根据需要插入的数据添加新行
		int len = dataList.size() - 1;
		for (int i=0; i<len; i++) {
			createRow(table, table.getRow(index), (i+1+index));
		}
		
		// 判断是否是自定义表格
		if (index==0) {
			// 根据每一行数据的数组大小，添加列数
			for (int i=0; i<dataList.size(); i++) {
				XWPFTableRow row = table.getRow(i);          // 每一行
				int cellSize = row.getTableCells().size();   // 每一行目前的列数
				int dataSize = dataList.get(i).length;       // 传入的List数据的每一个数据的数组大小
				
				// 给该行添加新列
				int colNum = dataSize - cellSize;
				if (colNum>0) {
					// 获取当前行最后一列的列属性
					XWPFTableCell sourceCell = row.getTableCells().get(cellSize-1);
					// 记录水平对齐方式
					String cellHorizontalAlignment = getCellTextAlign(sourceCell);
					// 记录垂直对齐方式
					String cellVerticalAlignment = getCellVertAlign(sourceCell);
					
					for (int j=0; j<colNum; j++) {
						XWPFTableCell addCell = row.addNewTableCell();
						addCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
						
						// 设置水平对齐方式
						if (cellHorizontalAlignment!=null && cellHorizontalAlignment.length()>0) {
							setCellHorizontalAlign(addCell, cellHorizontalAlignment);
						}
						// 设置垂直对齐方式
						if (cellVerticalAlignment!=null && cellVerticalAlignment.length()>0) {
							setCellVertAlign(addCell, cellVerticalAlignment);
						}
					}
				}
			}
		}
		
		// 记录每一列单元格的水平对齐方式
		List<String> horizontalAlignmentList = new ArrayList<String>();
		// 记录每一列单元格的垂直对齐方式
		List<String> verticalAlignmentList = new ArrayList<String>();
		
		// 遍历表格插入数据
		len = table.getRows().size();
		for (int i=index; i<len; i++) {
			XWPFTableRow newRow = table.getRow(i);
			List<XWPFTableCell> cellList = newRow.getTableCells();
			for (int j=0; j<cellList.size(); j++) {
				XWPFTableCell cell = cellList.get(j);
				String text = null;
				try {
					text = dataList.get(i-index)[j];
				} catch (Exception e) {
					
				} finally {
					if (text==null) {
						text = "";
					}
					else if ("BLANK_LINE".equals(text)) {
						for (XWPFTableCell xwpfTableCell : cellList) {
							xwpfTableCell.getCTTc().getTcPr().addNewTcBorders().addNewLeft().setVal(STBorder.NIL);
							xwpfTableCell.getCTTc().getTcPr().addNewTcBorders().addNewRight().setVal(STBorder.NIL);
							xwpfTableCell.getCTTc().getTcPr().addNewTcBorders().addNewBottom().setVal(STBorder.NIL);
						}
						for (XWPFParagraph paragraph : cell.getParagraphs()) {
							super.replaceParagraph(paragraph, "");
						}
						continue;
					}
				}
				
				XWPFParagraph addParagraph = cell.getParagraphs().get(0);
				XWPFRun createRun = addParagraph.createRun();
				
				// 需要插入的文本
				String insertText = super.replaceBr(text);
				
				try {
					// 字体样式反序列化
					ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(text.getBytes("ISO-8859-1"));
					ObjectInputStream objectInputStream = new ObjectInputStream(byteArrayInputStream);
					Font fontStyle = (Font) objectInputStream.readObject();
					objectInputStream.close();
					byteArrayInputStream.close();
					
					// 设置字体样式
					runHelper.setFontStyle(createRun, fontStyle);
					
					// 需要插入的文本
					insertText = fontStyle.getText();
				} catch (Exception e) {
					
				} finally {
					runHelper.setText(createRun, insertText);
				}
				
				// 单元格水平、垂直对齐方式
				if (i==index) {
					// 记录水平对齐方式
					String cellHorizontalAlignment = getCellTextAlign(cell);
					
					if (cellHorizontalAlignment==null || cellHorizontalAlignment.length()==0) {
						horizontalAlignmentList.add(null);
						setCellHorizontalAlign(cell, "left");
					} else {
						horizontalAlignmentList.add(cellHorizontalAlignment);
					}
					
					// 记录垂直对齐方式
					String cellVerticalAlignment = getCellVertAlign(cell);
					if (cellVerticalAlignment==null || cellVerticalAlignment.length()==0) {
						verticalAlignmentList.add(null);
					} else {
						verticalAlignmentList.add(cellVerticalAlignment.toString());
					}
				} else {
					try {
						// 设置水平对齐方式
						setCellHorizontalAlign(cell, horizontalAlignmentList.get(j));
						// 设置垂直对齐方式
						setCellVertAlign(cell, verticalAlignmentList.get(j));
					} catch (Exception e) {
						
					}
				}
			}
		}
	}

	/**
	 * 在表格指定位置新增一行
	 * @param table 需要插入数据的表格
	 * @param sourceRow 复制的源行
	 * @param rowIndex 表格指定位置
	 */
	private void createRow(XWPFTable table, XWPFTableRow sourceRow, int rowIndex) {
		// 在表格指定位置新增一行
		XWPFTableRow targetRow = table.insertNewTableRow(rowIndex);
		
		// 复制行属性
		targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
		
		// 复制列属性
		List<XWPFTableCell> cellList = sourceRow.getTableCells();
		if (cellList!=null && cellList.isEmpty()==false) {
			XWPFTableCell targetCell = null;
			for (XWPFTableCell sourceCell : cellList) {
				targetCell = targetRow.addNewTableCell();
				targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
			}
		}
	}
	
	/**
	 * 获取单元格水平对齐方式
	 * @param cell
	 * @return
	 */
	private String getCellTextAlign(XWPFTableCell cell) {
		CTTc cttc = cell.getCTTc();
		CTP ctp = cttc.getPList().get(0);
		CTPPr ctppr = ctp.getPPr();
		if (ctppr==null) {
			ctppr = ctp.addNewPPr();
		}
		CTJc ctjc = ctppr.getJc();
		if (ctjc==null) {
			ctjc = ctppr.addNewJc();
		}
		
		Enum enumVal = ctjc.getVal();
		if (enumVal!=null) {
			return enumVal.toString().toLowerCase();
		}
		
		return null;
	}
	
	/**
	 * 设置单元格水平对齐方式
	 * @param cell
	 * @param horizontalAlign
	 */
	private void setCellHorizontalAlign(XWPFTableCell cell, String horizontalAlign) {
		if (horizontalAlign!=null && horizontalAlign.length()>0) {
			horizontalAlign = horizontalAlign.toLowerCase();
			
			CTTc cttc = cell.getCTTc();
			CTP ctp = cttc.getPList().get(0);
			CTPPr ctppr = ctp.getPPr();
			if (ctppr==null) {
				ctppr = ctp.addNewPPr();
			}
			CTJc ctjc = ctppr.getJc();
			if (ctjc==null) {
				ctjc = ctppr.addNewJc();
			}
			
			if (horizontalAlign.equals("left")) {
				ctjc.setVal(STJc.LEFT);
			}
			else if (horizontalAlign.equals("center")) {
				ctjc.setVal(STJc.CENTER);
			}
			else if (horizontalAlign.equals("right")) {
				ctjc.setVal(STJc.RIGHT);
			}
		}
	}
	
	/**
	 * 获取单元格垂直对齐方式
	 * @param cell
	 * @return
	 */
	private String getCellVertAlign(XWPFTableCell cell) {
		XWPFVertAlign verticalAlignment = cell.getVerticalAlignment();
		if (verticalAlignment!=null) {
			return verticalAlignment.toString().toLowerCase();
		}
		
		return null;
	}
	
	/**
	 * 设置单元格垂直对齐方式
	 * @param cell
	 * @param vertAlign
	 */
	private void setCellVertAlign(XWPFTableCell cell, String vertAlign) {
		if (vertAlign!=null && vertAlign.length()>0) {
			vertAlign = vertAlign.toLowerCase();
			
			if (vertAlign.equals("top")) {
				cell.setVerticalAlignment(XWPFVertAlign.TOP);
			}
			else if (vertAlign.equals("center")) {
				cell.setVerticalAlignment(XWPFVertAlign.CENTER);
			}
			else if (vertAlign.equals("both")) {
				cell.setVerticalAlignment(XWPFVertAlign.BOTH);
			}
			else if (vertAlign.equals("bottom")) {
				cell.setVerticalAlignment(XWPFVertAlign.BOTTOM);
			}
		}
	}
	
	/**
	 * 表格合并列
	 * @param table
	 * @param rowIndex 行的索引（从0开始计算）
	 * @param startCol 起始列（从0开始计算）
	 * @param endCol   终止列（从0开始计算）
	 */
	private void mergeCol(XWPFTable table, int rowIndex, int startCol, int endCol) {
		for (int i=startCol; i<=endCol; i++) {
			XWPFTableCell cell = table.getRow(rowIndex).getCell(i);
			if (i==startCol) {
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
			} else {
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	/**
	 * 表格合并行
	 * @param table
	 * @param colIndex 列的索引（从0开始计算）
	 * @param startRow 起始行（从0开始计算）
	 * @param endRow   终止行（从0开始计算）
	 */
	private void mergeRow(XWPFTable table, int colIndex, int startRow, int endRow) {
		for (int i=startRow; i<=endRow; i++) {
			XWPFTableCell cell = table.getRow(i).getCell(colIndex);
			if (i==startRow) {
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
			} else {
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
			}
		}
	}
	
	/**
	 * 正则匹配字符串
	 * @param str
	 * @return
	 */
	private Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}
	
}
