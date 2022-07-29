package com.apachepoi.template.lab;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
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

public class MockUpTemplateFormator {

	public String templateSource = "D:/Input/EMPLOYEEDETAIL.docx";

	public static XWPFDocument document;
	public String inputJSONSource = "D:/sampleJson3.json";
	public static Map<String, Object> dataConversion;
	public static Map<String, String> newMap;
	static String value;

	public static final String TEMPLATE_PREFIX = "${";
	public static final String TEMPLATE_SUFIX = "}";

//////////////////////Static  custom methods
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

	public static void main(String[] args) throws IOException, JSONException, JsonMappingException, JsonParseException {

		MockUpTemplateFormator jsonFileToJsonString2 = new MockUpTemplateFormator();

		jsonFileToJsonString2.cloneTemplateFromTable(document);

}

	public void cloneTemplateFromTable(XWPFDocument document) throws FileNotFoundException, IOException {

		// Load the source template
		document = new XWPFDocument(new FileInputStream(templateSource));
		XWPFTable tableFromTemplateSource;
		CTTbl cTTblTemplate;
		XWPFTable tableCopy;
		XWPFTable table;
		XWPFTableRow row;
		XWPFTableCell cell;
		XmlCursor cursor;
		XWPFParagraph paragraph;
		XWPFRun run;

		/**
		 * get first table (the template)
		 */
		tableFromTemplateSource = document.getTableArray(0);
		cTTblTemplate = tableFromTemplateSource.getCTTbl();
		cursor = setCursorToNextStartToken(cTTblTemplate);

		// Read the json data
		JSONArray jsonArray = convertDataTOJSONFromFile(inputJSONSource);

		if (jsonArray == null) {
			System.out.println("Unable to contninue invalid data");
			return;
		}

		System.out.println(jsonArray);

		for (int i = 0; i < jsonArray.length(); i++) {
			JSONObject jsonObject = jsonArray.getJSONObject(i);

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
