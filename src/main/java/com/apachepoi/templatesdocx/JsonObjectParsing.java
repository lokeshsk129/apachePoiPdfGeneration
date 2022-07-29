package com.apachepoi.templatesdocx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

import com.fasterxml.jackson.core.JsonParseException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;

public class JsonObjectParsing {

	private static final Logger LOGGER = Logger.getLogger("DocxTempalte.class");

	static String imgFile1 = "D:/card.jpg";
	static String resourcePath1 = "D:/Input/EMPLOYEEDETAIL.docx";
	static String docPath1 = "D:/firstDoc21.docx";
	static String pdfPath1 = "D:/NewPdf21.pdf";
    static String jsonFile = "D:/sampleJson3.json";
	

	public static void main(String[] args) throws Exception, JsonParseException, JsonMappingException {
		JsonObjectParsing JsonObjectParsing = new JsonObjectParsing();
		

		InputStream inputFile = new FileInputStream(resourcePath1);
		XWPFDocument xwpfDocument = new XWPFDocument(inputFile);

		ObjectMapper mapper = new ObjectMapper();

		File jsonfile = new File(jsonFile);	

		Map<String, Object> dataConversion = new HashMap<String, Object>();
		Map<String, String> newJsonMap = new HashMap<String, String>();
		for (Map.Entry<String, Object> entry : dataConversion.entrySet()) {
			if (entry.getValue() instanceof String) {
				newJsonMap.put(entry.getKey(), (String) entry.getValue());
			}
		}
		newJsonMap = mapper.readValue(jsonfile, new TypeReference<Map<String, String>>() {
		});

		

			replacePlaceholderInTables(newJsonMap, xwpfDocument);

			replaceImageInTables(newJsonMap, xwpfDocument, imgFile1);

			replacePlaceholdersInParagraphs(newJsonMap, xwpfDocument);

			saveWord(docPath1, xwpfDocument);
		

		JsonObjectParsing.ConvertToPDF(docPath1, pdfPath1);

		LOGGER.info(pdfPath1.toUpperCase());
		System.out.println("successfully converted pdf");

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

	private static void replacePlaceholderInTables(Map<String, String> newJsonMap, XWPFDocument xwpfDocument)
			throws Exception {
		for (Map.Entry<String, String> entry : newJsonMap.entrySet()) {
			for (XWPFTable xwpfTable : xwpfDocument.getTables()) {
				
				for (XWPFTableRow xwpfTableRow : xwpfTable.getRows()) {
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

	}

	static void replaceImageInTables(Map<String, String> newJsonMap, XWPFDocument xwpfDocument, String imgFile)
			throws Exception, NullPointerException {
		FileInputStream is = new FileInputStream(imgFile);
		for (Map.Entry<String, String> entry : newJsonMap.entrySet()) {
			for (XWPFTable xwpfTable1 : xwpfDocument.getTables()) {
				for (XWPFTableRow xwpfTableRow : xwpfTable1.getRows()) {
					for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
						for (XWPFParagraph xwpfParagraph1 : xwpfTableCell.getParagraphs()) {
							for (XWPFRun xwpfRun : xwpfParagraph1.getRuns()) {
								xwpfRun.getDocument();
								xwpfTable1.removeBorders();
								String text = xwpfRun.text();
								if (text != null && text.contains(entry.getKey()) && entry.getValue() != null
										&& !entry.getValue().isEmpty()) {
									text = text.replace("$" + entry.getKey() + "", "");
									xwpfRun.setText("", 0);
									xwpfRun.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(165),
											Units.toEMU(120));
									is.close();
								}
							}
						}
					}
				}
			}
		}
	}

	static void replacePlaceholdersInParagraphs(Map<String, String> newJsonMap, XWPFDocument xwpfDocument)
			throws Exception {
		for (Map.Entry<String, String> entry : newJsonMap.entrySet()) {
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

	static void changeOrientation(XWPFDocument xwpfDocument, String orientation) {
		CTDocument1 doc = xwpfDocument.getDocument();
		CTBody body = doc.getBody();
		CTSectPr section = body.addNewSectPr();
		XWPFParagraph para = xwpfDocument.createParagraph();
		CTP ctp = para.getCTP();
		CTPPr br = ctp.addNewPPr();
		br.setSectPr(section);
		CTPageSz pageSize = section.getPgSz();
		if (orientation.equals("landscape")) {
			pageSize.setOrient(STPageOrientation.LANDSCAPE);
			pageSize.setW(BigInteger.valueOf(842 * 20));
			pageSize.setH(BigInteger.valueOf(595 * 20));
		} else {
			pageSize.setOrient(STPageOrientation.PORTRAIT);
			pageSize.setH(BigInteger.valueOf(842 * 20));
			pageSize.setW(BigInteger.valueOf(595 * 20));
		}
	}

	
	private static void saveWord(String filepath, XWPFDocument xwpfDocument) throws Exception {
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(filepath);
			xwpfDocument.write(out);

		} catch (Exception e) {
			e.printStackTrace();
		} finally {

			out.close();
			out.flush();

		}

	}

}
