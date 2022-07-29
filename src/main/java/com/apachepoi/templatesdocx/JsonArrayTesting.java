package com.apachepoi.templatesdocx;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;


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

import com.apachepoi.template.lab.MockUpTemplateFormator;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.google.gson.JsonParseException;

public class JsonArrayTesting {

	public String templateSource = "D:/Input/EMPLOYEEDETAIL.docx";
	public String filepath = "D:/sampleJson3.json";
    public static XWPFDocument document;
    public static JSONArray jsonArray;
	public static	JSONObject jsonObject;
	static String value;

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
		while (cursor.hasNextToken() && cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START)
			;
		return cursor;
	}

	/*** @return the JSONObject and this method for file reader*/
	public JSONObject jsonFileLoader(String filename) throws JSONException, IOException {
		String content = new String(Files.readAllBytes(Paths.get(filename)));
		return new JSONObject(content);
	}

	public static void main(String[] args) throws IOException, JSONException, JsonMappingException, JsonParseException {

		MockUpTemplateFormator jsonFileToJsonString2 = new MockUpTemplateFormator();
		
        jsonFileToJsonString2.cloneTemplateFromTable(document);

        
        
	}

	public void cloneTemplateFromTable(XWPFDocument document) throws FileNotFoundException, IOException {

		/** Load the source template*/
		document = new XWPFDocument(new FileInputStream(templateSource));
		XWPFTable tableFromTemplateSource;
		CTTbl cTTblTemplate;
		XWPFTable tableCopy;
		XWPFTable table;
		XmlCursor cursor;
		XWPFParagraph paragraph;
		XWPFRun run;

		/** get first table (the template)*/
		tableFromTemplateSource = document.getTableArray(0);
		cTTblTemplate = tableFromTemplateSource.getCTTbl();
		cursor = setCursorToNextStartToken(cTTblTemplate);

		BufferedReader reader = new BufferedReader(new FileReader(filepath));
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
                } catch (Exception e) {
				System.out.println("Error while reading the data from Source file: " + e.getLocalizedMessage());
				}

			/** place the first data in the table[0] */
			for (int d = 0; d < jsonObject.length(); d++) {
				for (XWPFTableRow xwpfTableRow : tableFromTemplateSource.getRows()) {
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
			
			
			/** insert new empty paragraph */
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

				for (int d = 0; d < jsonObject.length(); d++) {
					for (XWPFTableRow xwpfTableRow : tableFromTemplateSource.getRows()) {
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
			}

			paragraph = document.insertNewParagraph(cursor);
			run = paragraph.createRun();
		    run.setText("");
			cursor = setCursorToNextStartToken(paragraph.getCTP());
		
		FileOutputStream out = new FileOutputStream("D:/Input/WordResult1.docx");
		document.write(out);
		out.close();
		document.close();
			}
	}
}
