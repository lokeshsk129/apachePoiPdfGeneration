package com.apachepoi.templatesdocx;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

public class DynamicObjectInsert {

	private static final Logger LOGGER = Logger.getLogger("DynamicObjectInsert.class");

	public static String docPath1 = "D:/Input/EDOut5.docx";
	public static String pdfPath1 = "D:/Input/EDOut5.pdf";
	public static String templateSource = "D:/Input/EMPLOYEEDETAIL.docx";
	public static String inputJSONSource = "D:/sampleJson3.json";
	public static String destinationFile = "D:/image.jpg"; 
	

	public static final String TEMPLATE_PREFIX = "${";
	public static final String TEMPLATE_SUFIX = "}";
	public static final String TEMPLATE_TEXT = "${avatar}";
	


	public static void main(String[] args) throws Exception {

	//	String destinationFile = "D:/image.jpg"; 

		new DynamicObjectInsert().construct(destinationFile);

		new DocxToPdfConversion().ConvertToPDF(docPath1, pdfPath1);

		LOGGER.info(pdfPath1);
	}
	
	
	
	static XmlCursor setCursorToNextStartToken(XmlObject object) {
		XmlCursor cursor = object.newCursor();
		cursor.toEndToken();
		while (cursor.hasNextToken() && cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START)
			;
		return cursor;
	}


	/** construct method for JSON Array to replace placeholder */
	private void construct(String destinationFile) throws Exception {
		JSONArray jsonArray = convertDataTOJSONFromFile(inputJSONSource);
		if (jsonArray == null) {
			System.out.println("Unable to contninue invalid data");
			return;
		}

		XWPFDocument document = new XWPFDocument(new FileInputStream(templateSource));
		XWPFTable tableCopy;
		XWPFParagraph paragraph;

		/** get first table (the template) */
		XWPFTable tableTemplate = document.getTableArray(0);
		CTTbl cTTblTemplate = tableTemplate.getCTTbl();
		XmlCursor cursor = setCursorToNextStartToken(cTTblTemplate);

		/** creating empty paragraph */
		paragraph = document.insertNewParagraph(cursor);
		cursor = setCursorToNextStartToken(paragraph.getCTP());

		for (int t = 0; t < jsonArray.length(); t++) {
			JSONObject jsonObject = jsonArray.getJSONObject(t);

			/** insert new empty table at position t */
			XWPFTable table2 = document.insertNewTbl(cursor);
			cursor = setCursorToNextStartToken(table2.getCTTbl());

			/** copy the template table */
			tableCopy = new XWPFTable((CTTbl) cTTblTemplate.copy(), document);

			replaceImageInTables(tableCopy, jsonObject, destinationFile);
			replaceTextInTables(tableCopy, jsonObject);

			/** set tableCopy at position t instead of table */
			document.setTable(t + 1, tableCopy);

			paragraph = document.insertNewParagraph(cursor);
			cursor = setCursorToNextStartToken(paragraph.getCTP());
			System.out.println(cursor.getTextValue());
		}

		deleteOneTable(document, 0);

		FileOutputStream out = new FileOutputStream("D:/Input/EDOut5.docx");
		document.write(out);
		out.close();
		document.close();

	}

	/** delete the table from template */
	private static void deleteOneTable(XWPFDocument document, int tableIndex) {
		try {
			int bodyElement = getBodyElementOfTable(document, tableIndex);
			document.removeBodyElement(bodyElement);
		} catch (Exception e) {
			System.out.println("There is no table #" + tableIndex + " in the document.");
		}
	}

	/** get the body element */
	private static int getBodyElementOfTable(XWPFDocument document, int tableNumberInDocument) {
		List<XWPFTable> tables = document.getTables();
		XWPFTable theTable = tables.get(tableNumberInDocument);
		return document.getPosOfTable(theTable);
	}

	/** replacing the text in placeholder cell */
	private void replaceTextInTables(XWPFTable table, JSONObject jsonObject) {
		for (XWPFTableRow xwpfTableRow : table.getRows()) {
			for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
				for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs()) {
					for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
						String text = xwpfRun.text();
						if (text.startsWith(TEMPLATE_PREFIX) && text.contains(TEMPLATE_SUFIX)) {
							text = text.substring(text.indexOf(TEMPLATE_PREFIX) + TEMPLATE_PREFIX.length(),
									text.indexOf(TEMPLATE_SUFIX));
							String value1 = (String) jsonObject.get(text);
							text = text.replace(text, value1);
							xwpfRun.setText(text, 0);
							setRun(xwpfRun, "Times New Roman", "", 16, true, false);

						}
					}
				}
			}
		}
	}

	/** setting text color,text font style and font size */
	private static XWPFRun setRun(XWPFRun xwpfRun, String fontFamily, String rgbColor, int fontSize, boolean bold,
			boolean addBreak) {
		xwpfRun.setFontFamily(fontFamily);
		xwpfRun.setFontSize(fontSize);
		xwpfRun.setColor(rgbColor);
		xwpfRun.setBold(bold);
		if (addBreak)
			xwpfRun.addBreak();
		return xwpfRun;
	}

	/** replacing the image in placeholder cell */
	static void replaceImageInTables(XWPFTable table, JSONObject jsonObject, String destinationFile)
			throws Exception, NullPointerException {
		FileInputStream is = new FileInputStream(destinationFile);
		for (XWPFTableRow xwpfTableRow : table.getRows()) {
			for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
				for (XWPFParagraph xwpfParagraph1 : xwpfTableCell.getParagraphs()) {
					xwpfParagraph1.setAlignment(ParagraphAlignment.CENTER);
					xwpfTableCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
					for (XWPFRun xwpfRun : xwpfParagraph1.getRuns()) {
						xwpfRun.getDocument();
						String text = xwpfRun.text();
						if (text.startsWith(TEMPLATE_PREFIX) && text.contains(TEMPLATE_TEXT)) {
							String url = jsonObject.get("avatar").toString();
							loadUrl(url, destinationFile);
							xwpfRun.setText("", 0);
							xwpfRun.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, destinationFile, Units.toEMU(120),
									Units.toEMU(110));

						}
					}
				}
			}
		}
	}

	public static String loadUrl(String url, String destinationFile) throws NullPointerException, Exception {

		URL url1 = new URL(url);
		InputStream is = url1.openStream();
		OutputStream os = new FileOutputStream(destinationFile);

		byte[] b = new byte[2048];
		int length;

		while ((length = is.read(b)) != -1) {
			os.write(b, 0, length);
		}
		is.close();
		os.close();
		return destinationFile;

	}

	/** replace text in paragraph */
	static void replacePlaceholdersInParagraphs(Map<String, String> data, XWPFDocument xwpfDocument) throws Exception {
		for (Map.Entry<String, String> entry : data.entrySet()) {
			for (XWPFParagraph paragraph : xwpfDocument.getParagraphs()) {
				for (XWPFRun run : paragraph.getRuns()) {
					String text = run.text();
					if (text != null && text.contains(entry.getKey()) && entry.getValue() != null
							&& !entry.getValue().isEmpty()) {
						text = text.replace(entry.getKey(), entry.getValue());
						run.setText(text, 0);

					}
				}
			}
		}
	}

	/** convert JSON File to JSON Array */
	public static JSONArray convertDataTOJSONFromFile(String filePath) {

		try {

			BufferedReader reader = new BufferedReader(new FileReader(filePath));
			StringBuilder stringBuilder = new StringBuilder();
			String line = null;
			String ls = System.getProperty("line.separator");
			while ((line = reader.readLine()) != null) {
				stringBuilder.append(line);
				stringBuilder.append(ls);
			}
			stringBuilder.deleteCharAt(stringBuilder.length() - 1);
			reader.close();
			String content = stringBuilder.toString();

			return new JSONArray(content);

		} catch (Exception e) {
			System.out.println("Error while reading the data from Source file: " + e.getLocalizedMessage());
			return null;
		}

	}

}