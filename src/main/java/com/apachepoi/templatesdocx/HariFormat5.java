package com.apachepoi.templatesdocx;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.List;

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

public class HariFormat5 {

	public String templateSource="D:/Input/EMPLOYEEDETAIL.docx";
	public String inputJSONSource="D:/sampleJson3.json";

	public static final String TEMPLATE_PREFIX = "${";
	public static final String TEMPLATE_SUFIX = "}";

	static XmlCursor setCursorToNextStartToken(XmlObject object) {

		XmlCursor cursor = object.newCursor();
		cursor.toEndToken(); // Now we are at end of the XmlObject.
		// There always must be a next start token.
		while (cursor.hasNextToken() && cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START)
			;
		// Now we are at the next start token and can insert new things here.
		return cursor;
	}

	static void removeCellValues(XWPFTableCell cell) {
		for (XWPFParagraph paragraph : cell.getParagraphs()) {
			for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
				paragraph.removeRun(i);
			}
		}
	}

	public static void main(String[] args) throws Exception {

		new HariFormat5().construct();
	}

	private void construct() throws Exception {

		JSONArray jsonArray = convertDataTOJSONFromFile(inputJSONSource);

		if (jsonArray == null) {
			System.out.println("Unable to contninue invalid data");
			return;
		}

		System.out.println(jsonArray);
		// The data. Each row a new table.

		XWPFDocument document = new XWPFDocument(new FileInputStream(templateSource));

		XWPFTable tableCopy;

		XWPFTable table;
		XWPFParagraph paragraph;
		XWPFRun run;

		// get first table (the template)
		XWPFTable tableTemplate = document.getTableArray(0);
		CTTbl cTTblTemplate = tableTemplate.getCTTbl();
		XmlCursor cursor = setCursorToNextStartToken(cTTblTemplate);

		System.out.println(cursor.getTextValue());

		// fill in first data in first table (the template)

		paragraph = document.insertNewParagraph(cursor);
		cursor = setCursorToNextStartToken(paragraph.getCTP());

		// fill in next data, each data row in one table
		for (int t = 0; t < jsonArray.length(); t++) {
			JSONObject jsonObject = jsonArray.getJSONObject(t);

			XWPFTable table2 = document.insertNewTbl(cursor); // insert new empty table at position t
			cursor = setCursorToNextStartToken(table2.getCTTbl());

			tableCopy = new XWPFTable((CTTbl) cTTblTemplate.copy(), document); // copy the template table

			// fill in data in tableCopy
			formatTableData(tableCopy, jsonObject);
			document.setTable(t + 1, tableCopy); // set tableCopy at position t instead of table

			paragraph = document.insertNewParagraph(cursor); // insert new empty paragraph
			cursor = setCursorToNextStartToken(paragraph.getCTP());
			System.out.println(cursor.getTextValue());
		}

		deleteOneTable(document, 0);

		FileOutputStream out = new FileOutputStream("D:/Input/EDOut5.docx");
		document.write(out);
		out.close();
		document.close();

	}

	private static void deleteOneTable(XWPFDocument document, int tableIndex) {
		try {
			int bodyElement = getBodyElementOfTable(document, tableIndex);
			document.removeBodyElement(bodyElement);
		} catch (Exception e) {
			System.out.println("There is no table #" + tableIndex + " in the document.");
		}
	}

	private static int getBodyElementOfTable(XWPFDocument document, int tableNumberInDocument) {
		List<XWPFTable> tables = document.getTables();
		XWPFTable theTable = tables.get(tableNumberInDocument);
		return document.getPosOfTable(theTable);
	}

	private void formatTableData(XWPFTable table, JSONObject jsonObject) {
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

						}

					}
				}
			}
		}
	}

	public JSONArray convertDataTOJSONFromFile(String filePath) {

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