package cn.javaex.officejj.word.help;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * 图片
 * 
 * @author 陈霓清
 */
public class ImageHelper {
	
	/**
	 * 根据图片类型，取得对应的图片类型代码
	 * @param imageType
	 * @return int
	 */
	public int getImageType(String imageType) {
		int type = XWPFDocument.PICTURE_TYPE_JPEG;
		
		if (imageType!=null && imageType.length()>0) {
			imageType = imageType.toLowerCase();
			
			switch (imageType) {
				case "png":
					type = XWPFDocument.PICTURE_TYPE_PNG;
					break;
				case "dib":
					type = XWPFDocument.PICTURE_TYPE_DIB;
					break;
				case "emf":
					type = XWPFDocument.PICTURE_TYPE_EMF;
					break;
				case "jpg":
					type = XWPFDocument.PICTURE_TYPE_JPEG;
					break;
				case "jpeg":
					type = XWPFDocument.PICTURE_TYPE_JPEG;
					break;
				case "wmf":
					type = XWPFDocument.PICTURE_TYPE_WMF;
					break;
				default:
					break;
			}
		}
		
		return type;
	}
	
}
