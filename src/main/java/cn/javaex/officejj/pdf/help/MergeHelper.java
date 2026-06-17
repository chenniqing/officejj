package cn.javaex.officejj.pdf.help;

import java.io.File;
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
		if (list==null || list.isEmpty()) {
			return;
		}
		if (destPath==null || destPath.trim().length()==0) {
			throw new IllegalArgumentException("PDF合并后的目标路径不能为空");
		}

		File targetFile = new File(destPath);
		File parentFile = targetFile.getParentFile();
		if (parentFile!=null && !parentFile.exists()) {
			parentFile.mkdirs();
		}

		PdfReader firstReader = null;
		Document document = null;
		try {
			firstReader = new PdfReader(list.get(0));
			document = new Document(firstReader.getPageSize(1));
		} finally {
			if (firstReader!=null) {
				firstReader.close();
			}
		}

		PdfCopy copy = null;
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(targetFile);
			copy = new PdfCopy(document, out);
			document.open();

			for (int i=0; i<list.size(); i++) {
				PdfReader reader = null;
				try {
					reader = new PdfReader(list.get(i));
					int totalPages = reader.getNumberOfPages();    // 获得总页码
					for (int j=1; j<=totalPages; j++) {
						document.newPage();
						PdfImportedPage page = copy.getImportedPage(reader, j);    // 从当前PDF，获取第j页
						copy.addPage(page);
					}
				} finally {
					if (reader!=null) {
						reader.close();
					}
				}
			}
		} finally {
			if (document!=null && document.isOpen()) {
				document.close();
			} else if (copy!=null) {
				copy.close();
			}
			if (out!=null) {
				out.close();
			}
		}
	}

}
