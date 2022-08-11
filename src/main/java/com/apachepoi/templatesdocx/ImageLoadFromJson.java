package com.apachepoi.templatesdocx;

import java.awt.BorderLayout;
import java.awt.Image;
import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLConnection;
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
import org.json.JSONArray;
import org.json.JSONObject;


public class ImageLoadFromJson {

	
	public static String docPath1 = "D:/Input/EDOut5.docx";
	public static String pdfPath1 = "D:/Input/EDOut5.pdf";
	public String templateSource = "D:/Input/EMPLOYEEDETAIL.docx";
	public static String inputJSONSource = "D:/testJsonFile.json";
	public static URL url;

	public static final String TEMPLATE_PREFIX = "${";
	public static final String TEMPLATE_SUFIX = "}";
	public static final String TEMPLATE_TEXT = "${avatar}";

	public static void main(String[] args) throws Exception {

		ImageLoadFromJson imageLoadFromJson = new ImageLoadFromJson();
		// imageLoadFromJson.construct();
		String imageUrl = "https://media.istockphoto.com/photos/taj-mahal-mausoleum-in-agra-picture-id1146517111?k=20&m=1146517111&s=612x612&w=0&h=vHWfu6TE0R5rG6DJkV42Jxr49aEsLN0ML-ihvtim8kk=";
		String destinationFile = "D:/image.jpg";

		imageLoadFromJson.saveImage(destinationFile);
		// show(imageUrl);

	}

	public void saveImage(String destinationFile) throws NullPointerException, Exception {
		JSONArray jsonArray = convertDataTOJSONFromFile(inputJSONSource);
		if (jsonArray == null) {
			System.out.println("Unable to contninue invalid data");
			return;
		}

		XWPFDocument document = new XWPFDocument(new FileInputStream(templateSource));
		

		for (int t = 0; t < jsonArray.length(); t++) {
			JSONObject jsonObject = jsonArray.getJSONObject(t);
			
			String url = jsonObject.get("avatar").toString();
			System.out.println(url);

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
			}
		replaceImageInTables(document, destinationFile);
		System.out.println("done");
		
		FileOutputStream out = new FileOutputStream("D:/Input/EdImageOut5.docx");
		document.write(out);
		System.out.println("docx file is generated");
		out.close();
		document.close();
	
			

		
	}

	static void replaceImageInTables(XWPFDocument document, String destinationFile)
			throws Exception, NullPointerException {
		FileInputStream is = new FileInputStream(destinationFile);
		for (XWPFTable xwpfTable1 : document.getTables()) {
			for (XWPFTableRow xwpfTableRow : xwpfTable1.getRows()) {
				for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
					for (XWPFParagraph xwpfParagraph1 : xwpfTableCell.getParagraphs()) {
						xwpfParagraph1.setAlignment(ParagraphAlignment.CENTER);
						xwpfTableCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
						for (XWPFRun xwpfRun : xwpfParagraph1.getRuns()) {
							xwpfRun.getDocument(); 
							String text = xwpfRun.text();
							if (text.startsWith(TEMPLATE_PREFIX) && text.contains(TEMPLATE_TEXT)){
							xwpfRun.setText("", 0);
							xwpfRun.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, destinationFile, Units.toEMU(120),
									Units.toEMU(130));
							
							}
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