package com.apachepoi.templatesdocx;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

import com.fasterxml.jackson.databind.JsonMappingException;
import com.google.gson.JsonParseException;

public class JsonFileToJsonString {

	public String resourcePath1 = "D:/Input/EMPLOYEEDETAIL.docx";
	public static JSONArray jsonArray;
	public static JSONObject jsonObject;

	public String filename = "D:/sampleJson3.json";
	static String value;
	public static int t;
	public static XWPFTable tableCopy;
	public static XWPFRun run;

	public static final String TEMPLATE_PREFIX = "${";
	public static final String TEMPLATE_SUFIX = "}";
	
	
	static void removeCellValues(XWPFTableCell xwpfTableCell) {
		for (XWPFParagraph paragraph : xwpfTableCell.getParagraphs()) {
			for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
				paragraph.removeRun(i);
			}
		}
	}
	
	static XmlCursor setCursorToNextStartToken(XmlObject object) {
		XmlCursor cursor = object.newCursor();
		cursor.toEndToken();
		while (cursor.hasNextToken() && cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START);
		return cursor;
	}


	public static void main(String[] args) throws IOException, JSONException, JsonMappingException, JsonParseException {
		JsonFileToJsonString jsonFileToJsonString2 = new JsonFileToJsonString();

		jsonFileToJsonString2.cloneTemplateFromTable();

	}

	
	

	public void cloneTemplateFromTable() throws FileNotFoundException, IOException {

//		for (int i = 0; i < jsonArray.length(); i++) {
//			try {
//				jsonObject = jsonArray.getJSONObject(i);
//				System.out.println(jsonObject);
//
//			} catch (JSONException e) {
//				e.printStackTrace();
//				}

		BufferedReader reader = new BufferedReader(new FileReader(filename));
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
		jsonArray = new JSONArray(content);

		XWPFDocument document = new XWPFDocument(new FileInputStream(resourcePath1));
		XWPFTable tableTemplate;
		CTTbl cTTblTemplate;
		XWPFTable table;
		XmlCursor cursor;
		XWPFParagraph paragraph;

		/**
		 * get first table (the template)
		 */
		tableTemplate = document.getTableArray(0);
		cTTblTemplate = tableTemplate.getCTTbl();
		cursor = setCursorToNextStartToken(cTTblTemplate);

		/**
		 * insert new empty paragraph
		 */
		paragraph = document.insertNewParagraph(cursor);
		cursor = setCursorToNextStartToken(paragraph.getCTP());

		/**
		 * insert new empty table at position t fill in next data, each data row in one
		 * table copy the template table
		 */

		for (t = 0; t < jsonArray.length(); t++) {
			table = document.insertNewTbl(cursor);
			cursor = setCursorToNextStartToken(table.getCTTbl());
			tableCopy = new XWPFTable((CTTbl) cTTblTemplate.copy(), document);
		

				for (int d = 0; d<jsonArray.length(); d++) {
					for (XWPFTableRow xwpfTableRow : tableTemplate.getRows()) {
						for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
							for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs()) {
								for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
									String text = xwpfRun.text();
									if (text.startsWith(TEMPLATE_PREFIX) && text.contains(TEMPLATE_SUFIX)) {
										text = text.substring(text.indexOf(TEMPLATE_PREFIX) + TEMPLATE_PREFIX.length(),
										text.indexOf(TEMPLATE_SUFIX));
                                    	String value1 = (String) jsonObject.get(text);
										System.out.println(text + "-" + value1);
                                        text = text.replace(text, value1);
										xwpfRun.setText(text, 0);

									}
								}
							}
						}
					}
				}

		/** set tableCopy at position t instead of table */
		document.setTable(t, tableCopy);

		/** set tableCopy at position t instead of table */
		paragraph = document.insertNewParagraph(cursor);
		cursor = setCursorToNextStartToken(paragraph.getCTP());

		paragraph = document.insertNewParagraph(cursor);
		run = paragraph.createRun();
		// run.setText("");
		cursor = setCursorToNextStartToken(paragraph.getCTP());

		FileOutputStream out = new FileOutputStream("D:/Input/WordResult.docx");
		document.write(out);
		out.close();
		document.close();
		}
	}
}
