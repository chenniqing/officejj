package cn.javaex.officejj.pdf.help;

import java.io.FileOutputStream;
import java.util.List;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;

/**
 * 合并Pdf
 * 
 * @author 陈霓清
 */
public class MergeHelper {
	
	/**
	 * 合并PDF
	 * @param list          需要合并的文件的全路径list
	 * @param destPath      合并后的文件存储的全路径
	 * @throws Exception
	 */
	public void mergePdf(List<String> list, String destPath) throws Exception {
		Document document = new Document(new PdfReader(list.get(0)).getPageSize(1));
		PdfCopy copy = new PdfCopy(document, new FileOutputStream(destPath));
		
		document.open();
		
		for (int i=0; i<list.size(); i++) {
			PdfReader reader = new PdfReader(list.get(i));
			int totalPages = reader.getNumberOfPages();    // 获得总页码
			for (int j=1; j<=totalPages; j++) {
				document.newPage();
				PdfImportedPage page = copy.getImportedPage(reader, j);    // 从当前PDF，获取第j页
				copy.addPage(page);
			}
		}
		
		document.close();
	}
	
}
