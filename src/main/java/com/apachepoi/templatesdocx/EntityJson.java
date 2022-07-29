package com.apachepoi.templatesdocx;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

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


public class EntityJson {

	public String resourcePath1 = "D:/Input/EMPLOYEEDETAIL.docx";
	public static JSONArray jsonArray;
	public static JSONObject jsonObject;
	public static XWPFDocument document;
	public String filename = "D:/sampleJson3.json";
	public static Map<String, Object> dataConversion;
	public static Map<String, String> newMap;
	static String value;

	public static void main(String[] args) throws IOException, JSONException, JsonMappingException, JsonParseException {
		EntityJson jsonFileToJsonString2 = new EntityJson();

		jsonFileToJsonString2.cloneTemplateFromTable(document);

	}

	static void removeCellValues(XWPFTableCell xwpfTableCell) {
		for (XWPFParagraph paragraph : xwpfTableCell.getParagraphs()) {
			for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
				paragraph.removeRun(i);
			}
		}
	}

	/**
	 * 
	 * @param object
	 * @return cursor object
	 */
	static XmlCursor setCursorToNextStartToken(XmlObject object) {
		XmlCursor cursor = object.newCursor();
		cursor.toEndToken();
		while (cursor.hasNextToken() && cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START)
			;
		return cursor;
	}

	/**
	 * 
	 * @param filename
	 * @return the JSONObject and this method for file reader
	 * @throws JSONException
	 * @throws IOException
	 */
	public JSONObject jsonFileLoader(String filename) throws JSONException, IOException {
		String content = new String(Files.readAllBytes(Paths.get(filename)));
		return new JSONObject(content);
	}

	public void cloneTemplateFromTable(XWPFDocument document) throws FileNotFoundException, IOException {
		JsonArrayToObject jsonArrayToObject = new JsonArrayToObject();

		document = new XWPFDocument(new FileInputStream(resourcePath1));
		XWPFTable tableTemplate;
		CTTbl cTTblTemplate;
		XWPFTable tableCopy;
		XWPFTable table;
		XmlCursor cursor;
		XWPFParagraph paragraph;
		XWPFRun run;
		

		/**
		 * get first table (the template)
		 */
		tableTemplate = document.getTableArray(0);
		cTTblTemplate = tableTemplate.getCTTbl();
		cursor = setCursorToNextStartToken(cTTblTemplate);

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
	

		for (int i = 0; i < jsonArray.length(); i++) {
			try {
				jsonObject = jsonArray.getJSONObject(i);

				dataConversion = jsonArrayToObject.jsonTojavaMap(jsonObject);

				newMap = new HashMap<String, String>();
				for (Map.Entry<String, Object> entry : dataConversion.entrySet()) {
					if (entry.getValue() instanceof String) {
						newMap.put(entry.getKey(), (String) entry.getValue());
					}
				}

			} catch (JSONException e) {
				e.printStackTrace();
			}

			for (int c = 0; c < jsonArray.length(); c++) {

				for (Map.Entry<String, String> entry : newMap.entrySet()) {

					for (XWPFTableRow xwpfTableRow : tableTemplate.getRows()) {
						for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
							for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs()) {
								for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
									String text = xwpfRun.text();
									System.out.println(xwpfRun.text());
									if (text != null && text.contains(entry.getKey()) && entry.getValue() != null
											&& !entry.getValue().isEmpty()) {
										text = text.replace("${" + entry.getKey() + "}", entry.getValue());
										xwpfRun.setText(text, 0);
										System.out.println(text);

									}
								}
							}
						}
					}

				}
			}

			/**
			 * insert new empty paragraph
			 */
			paragraph = document.insertNewParagraph(cursor);
			cursor = setCursorToNextStartToken(paragraph.getCTP());

			/**
			 * insert new empty table at position t fill in next data, each data row in one
			 * table copy the template table
			 */
			for (int t = 1; t < jsonArray.length(); t++) {
				table = document.insertNewTbl(cursor);
				cursor = setCursorToNextStartToken(table.getCTTbl());
				tableCopy = new XWPFTable((CTTbl) cTTblTemplate.copy(), document);

				for (int d = 0; d < jsonArray.length(); d++) {

					for (Map.Entry<String, String> entry : newMap.entrySet()) {

						for (XWPFTableRow xwpfTableRow : tableTemplate.getRows()) {
							for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
								for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs()) {
									for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
										String text = xwpfRun.text();
										System.out.println(xwpfRun.text());
										if (text != null && text.contains(entry.getKey()) && entry.getValue() != null
												&& !entry.getValue().isEmpty()) {
											text = text.replace("${" + entry.getKey() + "}", entry.getValue());
											xwpfRun.setText(text, 0);
											System.out.println(text);

										}
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
			}

			paragraph = document.insertNewParagraph(cursor);
			run = paragraph.createRun();
			run.setText("");
			cursor = setCursorToNextStartToken(paragraph.getCTP());
		
		FileOutputStream out = new FileOutputStream("D:/Input/WordResult.docx");
		document.write(out);
		out.close();
		document.close();
		}
	}

}
