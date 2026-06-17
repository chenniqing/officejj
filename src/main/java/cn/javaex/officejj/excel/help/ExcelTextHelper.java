package cn.javaex.officejj.excel.help;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import cn.javaex.officejj.excel.ExcelUtils;

/**
 * Excel文本格式辅助工具。
 * 提供CSV和简单HTML table互转能力，适合数据交换、预览片段和轻量导入导出。
 *
 * @author 陈霓清
 */
public class ExcelTextHelper {

	/**
	 * 读取CSV。
	 * 支持双引号包裹、双引号转义和逗号分隔。
	 * @param in 输入流
	 * @param charset 字符集
	 * @return
	 * @throws Exception
	 */
	public List<String[]> readCsv(InputStream in, Charset charset) throws Exception {
		if (in==null) {
			throw new IllegalArgumentException("CSV输入流不能为空");
		}
		if (charset==null) {
			charset = Charset.forName("UTF-8");
		}

		List<String[]> list = new ArrayList<String[]>();
		try (BufferedReader reader = new BufferedReader(new InputStreamReader(in, charset))) {
			String line;
			while ((line = reader.readLine())!=null) {
				list.add(this.parseCsvLine(line));
			}
		}
		return list;
	}

	/**
	 * 写出CSV。
	 * @param rowList 行数据
	 * @param out 输出流
	 * @param charset 字符集
	 * @throws Exception
	 */
	public void writeCsv(List<String[]> rowList, OutputStream out, Charset charset) throws Exception {
		if (out==null) {
			throw new IllegalArgumentException("CSV输出流不能为空");
		}
		if (charset==null) {
			charset = Charset.forName("UTF-8");
		}

		OutputStreamWriter writer = new OutputStreamWriter(out, charset);
		if (rowList!=null) {
			for (String[] row : rowList) {
				writer.write(this.toCsvLine(row));
				writer.write("\r\n");
			}
		}
		writer.flush();
	}

	/**
	 * 将Sheet转成简单HTML table。
	 * 该方法只输出单元格文本，不输出复杂样式，适合轻量预览和数据交换。
	 * @param sheet Sheet对象
	 * @return
	 */
	public String toHtml(Sheet sheet) {
		if (sheet==null) {
			throw new IllegalArgumentException("Sheet不能为空");
		}

		StringBuilder sb = new StringBuilder();
		sb.append("<table>");
		int lastRowNum = sheet.getLastRowNum();
		for (int i=0; i<=lastRowNum; i++) {
			Row row = sheet.getRow(i);
			sb.append("<tr>");
			if (row!=null) {
				short lastCellNum = row.getLastCellNum();
				for (int j=0; j<Math.max(0, lastCellNum); j++) {
					Cell cell = row.getCell(j);
					sb.append("<td>").append(this.escapeHtml(ExcelUtils.getCellValue(cell))).append("</td>");
				}
			}
			sb.append("</tr>");
		}
		sb.append("</table>");
		return sb.toString();
	}

	/**
	 * 将简单HTML table写入Workbook。
	 * 只解析 tr、td、th 和文本内容，复杂CSS样式不参与转换。
	 * @param html HTML table文本
	 * @param workbook 目标Workbook
	 * @return
	 */
	public Workbook htmlToWorkbook(String html, Workbook workbook) {
		if (html==null || html.trim().length()==0) {
			throw new IllegalArgumentException("HTML内容不能为空");
		}
		if (workbook==null) {
			workbook = ExcelUtils.createWorkbook();
		}

		Sheet sheet = workbook.createSheet("Sheet1");
		Pattern rowPattern = Pattern.compile("(?is)<tr[^>]*>(.*?)</tr>");
		Pattern cellPattern = Pattern.compile("(?is)<t[dh][^>]*>(.*?)</t[dh]>");
		Matcher rowMatcher = rowPattern.matcher(html);
		int rowIndex = 0;
		while (rowMatcher.find()) {
			Row row = sheet.createRow(rowIndex++);
			Matcher cellMatcher = cellPattern.matcher(rowMatcher.group(1));
			int colIndex = 0;
			while (cellMatcher.find()) {
				String text = this.unescapeHtml(cellMatcher.group(1).replaceAll("(?is)<[^>]+>", ""));
				row.createCell(colIndex++).setCellValue(text);
			}
		}
		return workbook;
	}

	/**
	 * 解析CSV单行。
	 * @param line 行文本
	 * @return
	 */
	private String[] parseCsvLine(String line) {
		List<String> values = new ArrayList<String>();
		StringBuilder sb = new StringBuilder();
		boolean inQuote = false;
		for (int i=0; i<line.length(); i++) {
			char ch = line.charAt(i);
			if (ch=='"') {
				if (inQuote && i+1<line.length() && line.charAt(i+1)=='"') {
					sb.append('"');
					i++;
				} else {
					inQuote = !inQuote;
				}
			} else if (ch==',' && !inQuote) {
				values.add(sb.toString());
				sb.setLength(0);
			} else {
				sb.append(ch);
			}
		}
		values.add(sb.toString());
		return values.toArray(new String[values.size()]);
	}

	/**
	 * 转成CSV单行文本。
	 * @param row 行数据
	 * @return
	 */
	private String toCsvLine(String[] row) {
		if (row==null || row.length==0) {
			return "";
		}
		StringBuilder sb = new StringBuilder();
		for (int i=0; i<row.length; i++) {
			if (i>0) {
				sb.append(',');
			}
			String value = row[i]==null ? "" : row[i];
			boolean quote = value.indexOf(',')>=0 || value.indexOf('"')>=0 || value.indexOf('\n')>=0 || value.indexOf('\r')>=0;
			if (quote) {
				sb.append('"').append(value.replace("\"", "\"\"")).append('"');
			} else {
				sb.append(value);
			}
		}
		return sb.toString();
	}

	private String escapeHtml(String text) {
		if (text==null) {
			return "";
		}
		return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;");
	}

	private String unescapeHtml(String text) {
		if (text==null) {
			return "";
		}
		return text.replace("&quot;", "\"").replace("&gt;", ">").replace("&lt;", "<").replace("&amp;", "&").trim();
	}
}
