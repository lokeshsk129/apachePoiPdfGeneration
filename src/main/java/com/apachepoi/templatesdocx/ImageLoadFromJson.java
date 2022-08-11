package com.apachepoi.templatesdocx;

import java.awt.BorderLayout;
import java.awt.Image;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.util.List;

import javax.imageio.ImageIO;
import javax.swing.ImageIcon;
import javax.swing.JFrame;
import javax.swing.JLabel;

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

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;

public class ImageLoadFromJson {

	public static String docPath1 = "D:/Input/EdImageOut5.docx";
	public static String pdfPath1 = "D:/Input/EDOut5.pdf";
	public String templateSource = "D:/Input/EMPLOYEEDETAIL.docx";
	public static String inputJSONSource = "D:/testJsonFile.json";
	public static URL url;
	public static XWPFDocument document;

	public static final String TEMPLATE_PREFIX = "${";
	public static final String TEMPLATE_SUFIX = "}";
	public static final String TEMPLATE_TEXT = "${avatar}";

	static XmlCursor setCursorToNextStartToken(XmlObject object) {
		XmlCursor cursor = object.newCursor();
		cursor.toEndToken();
		while (cursor.hasNextToken() && cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START)
			;
		return cursor;
	}

	public static void main(String[] args) throws Exception {

		ImageLoadFromJson imageLoadFromJson = new ImageLoadFromJson();

		String imageUrl = "https://media.istockphoto.com/photos/taj-mahal-mausoleum-in-agra-picture-id1146517111?k=20&m=1146517111&s=612x612&w=0&h=vHWfu6TE0R5rG6DJkV42Jxr49aEsLN0ML-ihvtim8kk=";
		String destinationFile = "D:/image.jpg";

		imageLoadFromJson.saveImage(destinationFile);
		imageLoadFromJson.ConvertToPDF(docPath1, pdfPath1);

		
		
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

	public void saveImage(String destinationFile) throws NullPointerException, Exception {
		JSONArray jsonArray = convertDataTOJSONFromFile(inputJSONSource);
		if (jsonArray == null) {
			System.out.println("Unable to contninue invalid data");
			return;
		}

		document = new XWPFDocument(new FileInputStream(templateSource));
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

//			String url = jsonObject.get("avatar").toString();
//			System.out.println(url);
//			loadUrl(url, destinationFile);
			/** insert new empty table at position t */
			XWPFTable table2 = document.insertNewTbl(cursor);
			cursor = setCursorToNextStartToken(table2.getCTTbl());

			/** copy the template table */
			tableCopy = new XWPFTable((CTTbl) cTTblTemplate.copy(), document);


			replaceImageInTables(tableCopy, jsonObject, destinationFile);

			/** set tableCopy at position t instead of table */
			document.setTable(t + 1, tableCopy);

			paragraph = document.insertNewParagraph(cursor);
			cursor = setCursorToNextStartToken(paragraph.getCTP());
			System.out.println(cursor.getTextValue());
		}

		deleteOneTable(document, 0);

		FileOutputStream out = new FileOutputStream("D:/Input/EdImageOut5.docx");
		document.write(out);
		System.out.println("docx file is generated");
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
							System.out.println(url);
							loadUrl(url, destinationFile);
							xwpfRun.setText("", 0);
							xwpfRun.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, destinationFile, Units.toEMU(150),
									Units.toEMU(140));

						}
					}
				}
			}
		}
	}

	/** replacing the placeholder in cell */
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

	/** convert JSON File to JSON Array */
	public static void show(String urlLocation) {
		Image image = null;
		try {
			URL url = new URL(urlLocation);
			URLConnection conn = url.openConnection();
			conn.setRequestProperty("User-Agent", "Mozilla/5.0");

			conn.connect();
			InputStream urlStream = conn.getInputStream();
			image = ImageIO.read(urlStream);

			JFrame frame = new JFrame();
			JLabel lblimage = new JLabel(new ImageIcon(image));
			frame.getContentPane().add(lblimage, BorderLayout.CENTER);
			frame.setSize(image.getWidth(null) + 30, image.getHeight(null) + 30);
			frame.setVisible(true);

		} catch (IOException e) {
			System.out.println("Something went wrong, sorry:" + e.toString());
			e.printStackTrace();
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
//		replaceImageInTables(document, destinationFile);
//		System.out.println("done");
		return destinationFile;

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