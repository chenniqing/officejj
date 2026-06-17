package cn.javaex.officejj.word.help;

import java.io.IOException;
import java.math.BigInteger;

import javax.xml.namespace.QName;

import com.microsoft.schemas.vml.CTShape;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFldChar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;

import cn.javaex.officejj.common.entity.Picture;

/**
 * word文档
 * 
 * @author 陈霓清
 */
public class WordHelper extends Helper {

	private static final QName VML_SHAPE = new QName("urn:schemas-microsoft-com:vml", "shape");
	private static final STHdrFtr.Enum[] WATERMARK_HEADER_TYPES = new STHdrFtr.Enum[] {
			XWPFHeaderFooterPolicy.DEFAULT,
			XWPFHeaderFooterPolicy.FIRST,
			XWPFHeaderFooterPolicy.EVEN
	};
	
	/**
	 * 创建段落数组
	 * @param word
	 * @param obj
	 * @param align
	 * @return
	 */
	private XWPFParagraph[] createParagraphs(XWPFDocument word, Object obj, ParagraphAlignment align) {
		XWPFParagraph paragraph = new XWPFParagraph(CTP.Factory.newInstance(), word);
		paragraph.setAlignment(align);
		paragraph.setVerticalAlignment(TextAlignment.CENTER);
		
		RunHelper runHelper = new RunHelper();
		
		XWPFRun run = paragraph.createRun();
		runHelper.setValue(run, obj);
		
		XWPFParagraph[] paragraphs = new XWPFParagraph[1];
		paragraphs[0] = paragraph;
		
		return paragraphs;
	}

	/**
	 * 创建段落数组
	 * @param word
	 * @param obj1
	 * @param obj2
	 * @param spacing 
	 * @return
	 */
	private XWPFParagraph[] createParagraphs(XWPFDocument word, Object obj1, Object obj2, int spacing) {
		XWPFParagraph paragraph = new XWPFParagraph(CTP.Factory.newInstance(), word);
		paragraph.setAlignment(ParagraphAlignment.LEFT);
		paragraph.setVerticalAlignment(TextAlignment.CENTER);
		
		CTTabStop tabStop = paragraph.getCTP().getPPr().addNewTabs().addNewTab();
		tabStop.setVal(STTabJc.RIGHT);
		tabStop.setPos(BigInteger.valueOf(spacing * 6));
		
		RunHelper runHelper = new RunHelper();
		
		XWPFRun run = paragraph.createRun();
		runHelper.setValue(run, obj1);
		run.addTab();
		runHelper.setValue(run, obj2);
		
		XWPFParagraph[] paragraphs = new XWPFParagraph[1];
		paragraphs[0] = paragraph;
		
		return paragraphs;
	}
	
	/**
	 * 创建 header-footer
	 * @param word
	 * @param content
	 * @return
	 */
	private XWPFHeaderFooterPolicy createHeaderFooter(XWPFDocument word, String content) {
		XWPFHeaderFooterPolicy headerFooterPolicy = word.getHeaderFooterPolicy();
		if (headerFooterPolicy==null) {
			headerFooterPolicy = word.createHeaderFooterPolicy();
		}
		// 添加水印内容
		headerFooterPolicy.createWatermark(content);
		
		return headerFooterPolicy;
	}
	
	/**
	 * 设置水印
	 * @param word
	 * @param content     水印文字内容
	 */
	public void setWatermark(XWPFDocument word, String content) {
		// 创建 header-footer
		XWPFHeaderFooterPolicy headerFooterPolicy = this.createHeaderFooter(word, content);
		this.ensureWatermarkParagraphs(headerFooterPolicy, content);
		this.applyWatermarkStyle(headerFooterPolicy);
	}

	/**
	 * 补齐已有页眉中的水印段落。
	 * <p>POI 的 createWatermark 遇到已存在的页眉会直接复用，不会把水印段落追加进去。</p>
	 * @param headerFooterPolicy 页眉页脚策略
	 * @param content 水印文字内容
	 */
	private void ensureWatermarkParagraphs(XWPFHeaderFooterPolicy headerFooterPolicy, String content) {
		XWPFDocument sourceWord = null;
		try {
			XWPFHeaderFooterPolicy sourcePolicy = null;
			for (STHdrFtr.Enum type : WATERMARK_HEADER_TYPES) {
				XWPFHeader targetHeader = headerFooterPolicy.getHeader(type);
				if (targetHeader!=null && !this.hasWatermarkShape(targetHeader)) {
					if (sourcePolicy==null) {
						sourceWord = new XWPFDocument();
						sourcePolicy = sourceWord.createHeaderFooterPolicy();
						sourcePolicy.createWatermark(content);
					}
					this.copyWatermarkParagraph(sourcePolicy.getHeader(type), targetHeader);
				}
			}
		} finally {
			if (sourceWord!=null) {
				try {
					sourceWord.close();
				} catch (IOException e) {
					throw new IllegalStateException("关闭临时水印文档失败", e);
				}
			}
		}
	}

	/**
	 * 复制水印段落到目标页眉。
	 * @param sourceHeader 源页眉
	 * @param targetHeader 目标页眉
	 */
	private void copyWatermarkParagraph(XWPFHeader sourceHeader, XWPFHeader targetHeader) {
		if (sourceHeader==null || sourceHeader.getParagraphArray(0)==null) {
			return;
		}
		XWPFParagraph paragraph = targetHeader.createParagraph();
		paragraph.getCTP().set(sourceHeader.getParagraphArray(0).getCTP().copy());
	}

	/**
	 * 设置水印样式。
	 * <p>文档可能已经存在普通页眉，水印形状不一定在默认页眉的第一个 run 中，需要遍历所有页眉中的 VML 图形。</p>
	 * @param headerFooterPolicy 页眉页脚策略
	 */
	private void applyWatermarkStyle(XWPFHeaderFooterPolicy headerFooterPolicy) {
		for (STHdrFtr.Enum type : WATERMARK_HEADER_TYPES) {
			XWPFHeader header = headerFooterPolicy.getHeader(type);
			if (header!=null) {
				this.applyWatermarkStyle(header);
			}
		}
	}

	/**
	 * 设置页眉内水印样式。
	 * @param header 页眉
	 */
	private void applyWatermarkStyle(XWPFHeader header) {
		for (XWPFParagraph paragraph : header.getParagraphs()) {
			for (int i = 0; i < paragraph.getCTP().sizeOfRArray(); i++) {
				CTR ctr = paragraph.getCTP().getRArray(i);
				for (int j = 0; j < ctr.sizeOfPictArray(); j++) {
					XmlObject[] xmlobjects = ctr.getPictArray(j).selectChildren(VML_SHAPE);
					this.applyWatermarkShapeStyle(xmlobjects);
				}
			}
		}
	}

	/**
	 * 设置水印形状样式。
	 * @param xmlobjects VML shape 节点集合
	 */
	private void applyWatermarkShapeStyle(XmlObject[] xmlobjects) {
		for (XmlObject xmlobject : xmlobjects) {
			CTShape ctshape = (CTShape)xmlobject;
			if (!this.isWatermarkShape(ctshape)) {
				continue;
			}
			ctshape.setFillcolor("#EEEEEE");
			if (ctshape.getStyle()==null || !ctshape.getStyle().contains("rotation:315")) {
				ctshape.setStyle((ctshape.getStyle()==null ? "" : ctshape.getStyle()) + ";rotation:315");
			}
		}
	}

	/**
	 * 判断页眉中是否已经存在水印形状。
	 * @param header 页眉
	 * @return true 表示已存在水印
	 */
	private boolean hasWatermarkShape(XWPFHeader header) {
		for (XWPFParagraph paragraph : header.getParagraphs()) {
			for (int i = 0; i < paragraph.getCTP().sizeOfRArray(); i++) {
				CTR ctr = paragraph.getCTP().getRArray(i);
				for (int j = 0; j < ctr.sizeOfPictArray(); j++) {
					XmlObject[] xmlobjects = ctr.getPictArray(j).selectChildren(VML_SHAPE);
					for (XmlObject xmlobject : xmlobjects) {
						if (this.isWatermarkShape((CTShape)xmlobject)) {
							return true;
						}
					}
				}
			}
		}
		return false;
	}

	/**
	 * 判断 VML shape 是否为 POI 生成的水印形状。
	 * @param ctshape VML shape
	 * @return true 表示水印形状
	 */
	private boolean isWatermarkShape(CTShape ctshape) {
		return ctshape.getId()!=null && ctshape.getId().startsWith("PowerPlusWaterMarkObject");
	}

	/**
	 * 设置页眉
	 * @param word
	 * @param obj    字符串
	 *               cn.javaex.officejj.common.entity.Font
	 *               cn.javaex.officejj.common.entity.Picture
	 * @param align
	 */
	public void setHeader(XWPFDocument word, Object obj, ParagraphAlignment align) {
		if (obj instanceof Picture) {
			// 创建 header-footer
			XWPFHeaderFooterPolicy headerFooterPolicy = this.createHeaderFooter(word, "");
			
			// 获取默认标题，但此代码不会更新其他头
			XWPFHeader header = headerFooterPolicy.getHeader(XWPFHeaderFooterPolicy.DEFAULT);
			XWPFParagraph paragraph = header.getParagraphArray(0);
			paragraph.setAlignment(align);
			paragraph.setVerticalAlignment(TextAlignment.CENTER);
			
			RunHelper runHelper = new RunHelper();
			
			XWPFRun run = paragraph.createRun();
			runHelper.setValue(run, obj);
		} else {
			// 创建段落数组
			XWPFParagraph[] paragraphs = this.createParagraphs(word, obj, align);
			
			// 创建 header-footer
			CTSectPr sectPr = word.getDocument().getBody().addNewSectPr();
			XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(word, sectPr);
			headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, paragraphs);
		}
	}
	
	/**
	 * 设置页眉
	 * @param word
	 * @param obj1    字符串
	 *                cn.javaex.officejj.common.entity.Font
	 *                cn.javaex.officejj.common.entity.Picture
	 * @param obj2    字符串
	 *                cn.javaex.officejj.common.entity.Font
	 *                cn.javaex.officejj.common.entity.Picture
	 * @param spacing  间距，默认word下请填写 1440
	 */
	public void setHeader(XWPFDocument word, Object obj1, Object obj2, int spacing) {
		// 创建段落数组
		XWPFParagraph[] paragraphs = this.createParagraphs(word, obj1, obj2, spacing);
		
		// 创建 header-footer
		CTSectPr sectPr = word.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(word, sectPr);
		headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, paragraphs);
	}

	/**
	 * 设置页脚
	 * @param word
	 * @param obj    字符串
	 *               cn.javaex.officejj.common.entity.Font
	 *               cn.javaex.officejj.common.entity.Picture
	 * @param align
	 */
	public void setFooter(XWPFDocument word, Object obj, ParagraphAlignment align) {
		// 创建段落数组
		XWPFParagraph[] paragraphs = this.createParagraphs(word, obj, align);
		
		// 创建 header-footer
		CTSectPr sectPr = word.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(word, sectPr);
		headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, paragraphs);
	}

	/**
	 * 设置页脚
	 * @param word
	 * @param obj1    字符串
	 *                cn.javaex.officejj.common.entity.Font
	 *                cn.javaex.officejj.common.entity.Picture
	 * @param obj2    字符串
	 *                cn.javaex.officejj.common.entity.Font
	 *                cn.javaex.officejj.common.entity.Picture
	 * @param spacing  间距，默认word下请填写 1440
	 */
	public void setFooter(XWPFDocument word, Object obj1, Object obj2, int spacing) {
		// 创建段落数组
		XWPFParagraph[] paragraphs = this.createParagraphs(word, obj1, obj2, spacing);
		
		// 创建 header-footer
		CTSectPr sectPr = word.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(word, sectPr);
		headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, paragraphs);
	}

	/**
	 * 设置页码
	 * @param word
	 * @param align
	 */
	public void setPageNumber(XWPFDocument word, ParagraphAlignment align) {
		CTP ctp = CTP.Factory.newInstance();
		XWPFParagraph paragraph = new XWPFParagraph(ctp, word);
		XWPFRun run = paragraph.createRun();
		run.setText("第");
		run.setFontSize(11);
		
		run = paragraph.createRun();
		CTFldChar fldChar = run.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));
		
		run = paragraph.createRun();
		CTText ctText = run.getCTR().addNewInstrText();
		ctText.setStringValue("PAGE  \\* MERGEFORMAT");
		ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
		run.setFontSize(11);
		
		fldChar = run.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("end"));
		
		run = paragraph.createRun();
		run.setText("页 共");
		run.setFontSize(11);
		
		run = paragraph.createRun();
		fldChar = run.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));
		
		run = paragraph.createRun();
		ctText = run.getCTR().addNewInstrText();
		ctText.setStringValue("NUMPAGES  \\* MERGEFORMAT");
		ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
		run.setFontSize(11);
		
		fldChar = run.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("end"));
		
		run = paragraph.createRun();
		run.setText("页");
		run.setFontSize(11);
		
		paragraph.setAlignment(align);
		paragraph.setVerticalAlignment(TextAlignment.CENTER);
		XWPFParagraph[] paragraphs = new XWPFParagraph[1];
		paragraphs[0] = paragraph;
		CTSectPr sectPr = word.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(word, sectPr);
		headerFooterPolicy.createFooter(STHdrFtr.DEFAULT, paragraphs);
	}

}
