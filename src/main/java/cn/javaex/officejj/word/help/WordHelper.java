package cn.javaex.officejj.word.help;

import java.math.BigInteger;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFldChar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
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
		
		// 获取默认标题，但此代码不会更新其他头
		XWPFHeader header = headerFooterPolicy.getHeader(XWPFHeaderFooterPolicy.DEFAULT);
		XWPFParagraph paragraph = header.getParagraphArray(0);
		
		// 获取设置填充颜色和旋转的com.microsoft.schemas.vml.CTShape
		org.apache.xmlbeans.XmlObject[] xmlobjects = paragraph.getCTP()
				.getRArray(0)
				.getPictArray(0)
				.selectChildren(new javax.xml.namespace.QName("urn:schemas-microsoft-com:vml", "shape"));
		
		if (xmlobjects.length>0) {
			com.microsoft.schemas.vml.CTShape ctshape = (com.microsoft.schemas.vml.CTShape)xmlobjects[0];
			ctshape.setFillcolor("#EEEEEE");
			ctshape.setStyle(ctshape.getStyle() + ";rotation:315");
		}
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
