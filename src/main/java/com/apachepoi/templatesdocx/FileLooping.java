package com.apachepoi.templatesdocx;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import java.util.Iterator;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.json.JSONObject;


public class FileLooping {

	static String imgFile1 = "D:/card.jpg";
	static String resourcePath1 = "D:/SourceFile.docx";
	static String docPath1 = "D:/firstDoc21.docx";
	static String pdfPath1 = "D:/NewPdf21.pdf";
	static String jsonFile = "D:/testJsonFile.json";
	

	public String templateSource = "D:/Input/EMPLOYEEDETAIL.docx";
	public String inputJSONSource = "D:/sampleJson3.json";
	public static JSONObject obj1;
	public static String content;

	public static final String TEMPLATE_PREFIX = "${";
	public static final String TEMPLATE_SUFIX = "}";

	static XmlCursor setCursorToNextStartToken(XmlObject object) {
		XmlCursor cursor = object.newCursor();
		cursor.toEndToken(); 
		while (cursor.hasNextToken() && cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START);
		return cursor;
	}

	public static void main(String[] args) throws IOException {
		
		

		
	

		InputStream inputFile = new FileInputStream(resourcePath1);

		try (XWPFDocument doc = new XWPFDocument(inputFile)) {
			Iterator<IBodyElement> iter = doc.getBodyElementsIterator();
			while(iter.hasNext()) {
				IBodyElement element = iter.next();
				if (element instanceof XWPFTable) {
					String tableValue = ((XWPFTable) element).getText();
					System.out.println(tableValue);
					

				} else if (element instanceof XWPFParagraph) {
					String paravalue = ((XWPFParagraph) element).getText();
					System.out.println(paravalue);
					continue;

				}
				
			}
		}
	}

}
