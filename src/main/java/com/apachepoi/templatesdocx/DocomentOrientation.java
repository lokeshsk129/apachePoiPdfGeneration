package com.apachepoi.templatesdocx;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

public class DocomentOrientation {

	public static String docPath1 = "D:/Input/EDOut5.docx";
	public static String outdocPath1 = "D:/Input/OutEDOut5.docx";

	public static void main(String[] args) throws FileNotFoundException, IOException {
		DocomentOrientation DocomentOrientation = new DocomentOrientation();
		
		InputStream inputFile = new FileInputStream(docPath1);
		
		XWPFDocument xwpfDocument = new XWPFDocument(inputFile);

		DocomentOrientation.changeOrientation(xwpfDocument, "landscape");

	}

	/**
	 * setting page LANDSCAPE or PORTRAIT
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public void changeOrientation(XWPFDocument xwpfDocument, String orientation)
			throws FileNotFoundException, IOException {
		
		CTDocument1 doc = xwpfDocument.getDocument();
		CTBody body = doc.getBody();
		CTSectPr section = body.addNewSectPr();
		XWPFParagraph para = xwpfDocument.createParagraph();
		CTP ctp = para.getCTP();
		CTPPr br = ctp.addNewPPr();
		br.setSectPr(section);
		CTPageSz pageSize;
		if (section.isSetPgSz()) {
			pageSize = section.getPgSz();
		} else {
			pageSize = section.addNewPgSz();
		}
		if (orientation.equals("landscape")) {
			pageSize.setOrient(STPageOrientation.LANDSCAPE);
			pageSize.setW(BigInteger.valueOf(942 * 20));
			pageSize.setH(BigInteger.valueOf(695 * 20));
			System.out.println("LANDSCAPE done");
			
		}else if (orientation.equals("portrait")) {
			pageSize.setOrient(STPageOrientation.PORTRAIT);
			pageSize.setH(BigInteger.valueOf(842 * 20));
			pageSize.setW(BigInteger.valueOf(595 * 20));
			System.out.println("PORTRAIT done");
			FileOutputStream out = new FileOutputStream(outdocPath1);
			xwpfDocument.write(out);
			out.close();
			xwpfDocument.close();

		}
	}
	
	
	
}
