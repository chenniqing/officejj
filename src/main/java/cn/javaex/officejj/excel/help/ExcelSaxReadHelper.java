package cn.javaex.officejj.excel.help;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import cn.javaex.officejj.excel.entity.ExcelSaxReadResult;
import cn.javaex.officejj.excel.entity.ExcelSaxReadSetting;
import cn.javaex.officejj.excel.entity.ExcelSaxRow;
import cn.javaex.officejj.excel.function.ExcelReadCancelChecker;
import cn.javaex.officejj.excel.function.ExcelReadProgressListener;
import cn.javaex.officejj.excel.function.ExcelSaxRowHandler;

/**
 * Excel xlsx低内存读取。
 * 该实现基于POI事件模型，不创建完整Workbook，适合导入任务、大文件导入、批量入库。
 *
 * @author 陈霓清
 */
public class ExcelSaxReadHelper {

	/**
	 * 读取xlsx文件。
	 * @param in 输入流
	 * @param setting 读取配置
	 * @param rowHandler 批次行处理器
	 * @param progressListener 进度监听器，允许为空
	 * @param cancelChecker 取消检查器，允许为空
	 * @return
	 * @throws Exception
	 */
	public ExcelSaxReadResult read(InputStream in, ExcelSaxReadSetting setting, ExcelSaxRowHandler rowHandler,
			ExcelReadProgressListener progressListener, ExcelReadCancelChecker cancelChecker) throws Exception {
		if (in==null) {
			throw new IllegalArgumentException("Excel输入流不能为空");
		}
		if (rowHandler==null) {
			throw new IllegalArgumentException("Excel行处理器不能为空");
		}
		if (setting==null) {
			setting = new ExcelSaxReadSetting();
		}
		if (setting.getSheetNum()<0) {
			throw new IllegalArgumentException("Sheet序号不能小于0：" + setting.getSheetNum());
		}
		if (setting.getStartRowNum()<=0) {
			throw new IllegalArgumentException("起始行号必须从1开始：" + setting.getStartRowNum());
		}
		if (setting.getBatchSize()<=0) {
			throw new IllegalArgumentException("批次大小必须大于0：" + setting.getBatchSize());
		}

		ExcelSaxReadResult result = new ExcelSaxReadResult();
		try (OPCPackage pkg = OPCPackage.open(in)) {
			XSSFReader reader = new XSSFReader(pkg);
			StylesTable styles = reader.getStylesTable();
			ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
			DataFormatter formatter = new DataFormatter();
			XSSFReader.SheetIterator iterator = (XSSFReader.SheetIterator) reader.getSheetsData();

			int sheetNum = 0;
			while (iterator.hasNext()) {
				try (InputStream sheetStream = iterator.next()) {
					sheetNum++;
					if (setting.getSheetNum()>0 && setting.getSheetNum()!=sheetNum) {
						continue;
					}
					if (cancelChecker!=null && cancelChecker.isCancelled()) {
						result.setCancelled(true);
						break;
					}

					SaxSheetHandler sheetHandler = new SaxSheetHandler(sheetNum, iterator.getSheetName(), setting, rowHandler, progressListener, cancelChecker, result);
					XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
					parser.setContentHandler(new XSSFSheetXMLHandler(styles, null, strings, sheetHandler, formatter, false));
					parser.parse(new InputSource(sheetStream));
					sheetHandler.flush();
					result.setSheetCount(result.getSheetCount() + 1);
				}
				if (setting.getSheetNum()>0 && setting.getSheetNum()==sheetNum) {
					break;
				}
				if (result.isCancelled()) {
					break;
				}
			}
		}
		return result;
	}

	/**
	 * 单个Sheet的SAX事件处理器。
	 * 负责把分散的单元格事件组装成完整行，并按批次回调给业务代码。
	 */
	private static class SaxSheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
		private final int sheetNum;
		private final String sheetName;
		private final ExcelSaxReadSetting setting;
		private final ExcelSaxRowHandler rowHandler;
		private final ExcelReadProgressListener progressListener;
		private final ExcelReadCancelChecker cancelChecker;
		private final ExcelSaxReadResult result;
		private final List<ExcelSaxRow> batchList = new ArrayList<ExcelSaxRow>();
		private List<String> currentCellList;
		private int currentRowNum;

		private SaxSheetHandler(int sheetNum, String sheetName, ExcelSaxReadSetting setting, ExcelSaxRowHandler rowHandler,
				ExcelReadProgressListener progressListener, ExcelReadCancelChecker cancelChecker, ExcelSaxReadResult result) {
			this.sheetNum = sheetNum;
			this.sheetName = sheetName;
			this.setting = setting;
			this.rowHandler = rowHandler;
			this.progressListener = progressListener;
			this.cancelChecker = cancelChecker;
			this.result = result;
		}

		@Override
		public void startRow(int rowNum) {
			this.currentRowNum = rowNum + 1;
			this.currentCellList = new ArrayList<String>();
		}

		@Override
		public void endRow(int rowNum) {
			if (this.currentRowNum<setting.getStartRowNum()) {
				return;
			}
			if (!setting.isReadEmptyRow() && this.isEmpty(currentCellList)) {
				return;
			}
			if (setting.getMaxRows()>0 && result.getRowCount()>=setting.getMaxRows()) {
				result.setCancelled(true);
				return;
			}
			if (cancelChecker!=null && cancelChecker.isCancelled()) {
				result.setCancelled(true);
				return;
			}

			ExcelSaxRow row = new ExcelSaxRow();
			row.setSheetNum(sheetNum);
			row.setSheetName(sheetName);
			row.setRowNum(this.currentRowNum);
			row.setCellList(currentCellList);
			batchList.add(row);
			result.setRowCount(result.getRowCount() + 1);

			if (progressListener!=null) {
				progressListener.onProgress(sheetNum, result.getRowCount());
			}
			if (batchList.size()>=setting.getBatchSize()) {
				this.flush();
			}
		}

		@Override
		public void cell(String cellReference, String formattedValue, XSSFComment comment) {
			int colIndex = this.getColumnIndex(cellReference);
			while (currentCellList.size()<=colIndex) {
				currentCellList.add("");
			}
			currentCellList.set(colIndex, formattedValue==null ? "" : formattedValue);
		}

		@Override
		public void headerFooter(String text, boolean isHeader, String tagName) {
			// 页眉页脚不是导入数据，忽略即可。
		}

		/**
		 * 立即回调当前批次数据。
		 */
		private void flush() {
			if (batchList.isEmpty()) {
				return;
			}
			try {
				rowHandler.handle(new ArrayList<ExcelSaxRow>(batchList));
				batchList.clear();
			} catch (Exception e) {
				throw new RuntimeException("Excel SAX读取批次处理失败", e);
			}
		}

		/**
		 * 判断行是否为空。
		 * @param cellList 单元格集合
		 * @return
		 */
		private boolean isEmpty(List<String> cellList) {
			if (cellList==null || cellList.isEmpty()) {
				return true;
			}
			for (String value : cellList) {
				if (value!=null && value.length()>0) {
					return false;
				}
			}
			return true;
		}

		/**
		 * 由A1、BC12等单元格引用计算列索引。
		 * @param cellReference 单元格引用
		 * @return
		 */
		private int getColumnIndex(String cellReference) {
			if (cellReference==null || cellReference.length()==0) {
				return currentCellList.size();
			}
			int col = 0;
			for (int i=0; i<cellReference.length(); i++) {
				char ch = cellReference.charAt(i);
				if (ch>='A' && ch<='Z') {
					col = col * 26 + (ch - 'A' + 1);
				} else if (ch>='a' && ch<='z') {
					col = col * 26 + (ch - 'a' + 1);
				} else {
					break;
				}
			}
			return Math.max(0, col - 1);
		}
	}
}
