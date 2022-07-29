package com.apachepoi.templatesdocx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;

public class DocxToPdfConversion {

	static String docPath1 = "D:/Input/EDOut5.docx";
	static String pdfPath1 = "D:/Input/EDOut5.pdf";

	public static void main(String[] args) throws Exception {
		
		DocxToPdfConversion conversionObject = new DocxToPdfConversion(); 
		
		//conversionObject.ConvertToPDF(docPath1,pdfPath1);

	}

	public void ConvertToPDF(String docPath1, String pdfPath1) throws Exception, NullPointerException {

		try {
			InputStream doc = new FileInputStream(new File(docPath1));
			long start = System.currentTimeMillis();
			XWPFDocument document = new XWPFDocument(doc);
			PdfOptions options = PdfOptions.create();
			OutputStream out = new FileOutputStream(new File(pdfPath1));
			PdfConverter.getInstance().convert(document, out, options);
			System.out.println("firstDoc21.docx was converted to a NewPdf21 file in :: "
					+ (System.currentTimeMillis() - start) + " milli seconds");

		} catch (Exception ex) {
			System.out.print(ex.getMessage());
		}
	}
}