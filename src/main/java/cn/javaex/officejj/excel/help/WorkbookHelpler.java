package cn.javaex.officejj.excel.help;

import java.util.UUID;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Workbook
 * 
 * @author 陈霓清
 */
public class WorkbookHelpler {
	/** 导出超过多少条数据时，使用SXSSFWorkbook */
	public static final int MAX_SIZE = 50000;
	
	/**
	 * 创建Workbook
	 * @param size 导出条数
	 * @return
	 */
	public Workbook createWorkbook(int size) {
		if (size >= MAX_SIZE) {
			return new SXSSFWorkbook(1000);
		}
		return new XSSFWorkbook();
	}

	/**
	 * 设置只读
	 * @param wb
	 * @param password
	 */
	public void setReadOnly(Workbook wb, String password) {
		if (password==null || password.length()==0) {
			password = UUID.randomUUID().toString().replace("-", "");
		}
		
		for (int i=0; i<wb.getNumberOfSheets(); i++) {
			wb.getSheetAt(i).protectSheet(password);
		}
	}
	
}
