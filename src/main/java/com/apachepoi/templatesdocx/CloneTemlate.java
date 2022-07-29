package com.apachepoi.templatesdocx;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

public class CloneTemlate {
	
	static XmlCursor setCursorToNextStartToken(XmlObject object) {
		  XmlCursor cursor = object.newCursor();
		  cursor.toEndToken(); //Now we are at end of the XmlObject.
		  //There always must be a next start token.
		  while(cursor.hasNextToken() && cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START);
		  //Now we are at the next start token and can insert new things here.
		  return cursor;
		 }

		 static void removeCellValues(XWPFTableCell cell) {
		  for (XWPFParagraph paragraph : cell.getParagraphs()) {
		   for (int i = paragraph.getRuns().size()-1; i >= 0; i--) {
		    paragraph.removeRun(i);
		   }  
		  }
		 }

		 public static void main(String[] args) throws Exception {

		  //The data. Each row a new table.
		  String[][] data= new String[][] {
		   new String[] {"", "", ""},
		   new String[] {"", "", ""},
		   new String[] {"", "", ""},
		   new String[] {"", "", ""}
		  };

		 
		  XWPFDocument document = new XWPFDocument(new FileInputStream("D:/Lab/EMPLOYEEDETAIL.docx"));
		  XWPFTable tableTemplate;
		  CTTbl cTTblTemplate;
		  XWPFTable tableCopy;
		  XWPFTable table;
		  XWPFTableRow row;
		  XWPFTableCell cell;
		  XmlCursor cursor;
		  XWPFParagraph paragraph;
		  XWPFRun run;
		  String value;

		  //get first table (the template)
		  tableTemplate = document.getTableArray(0);
		  cTTblTemplate = tableTemplate.getCTTbl();
		  cursor = setCursorToNextStartToken(cTTblTemplate);

		  //fill in first data in first table (the template)
		  for (int c = 0; c < data[0].length; c++) {
		   value = data[0][c];
		   row = tableTemplate.getRow(1);
		   cell = row.getCell(c);
//		   removeCellValues(cell);
//		   cell.setText(value);
		  }

		  paragraph = document.insertNewParagraph(cursor); //insert new empty paragraph
		  cursor = setCursorToNextStartToken(paragraph.getCTP());

		  //fill in next data, each data row in one table
		  for (int t = 1; t < data.length; t++) {
		   table = document.insertNewTbl(cursor); //insert new empty table at position t
		   cursor = setCursorToNextStartToken(table.getCTTbl());

		   tableCopy = new XWPFTable((CTTbl)cTTblTemplate.copy(), document); //copy the template table

		   //fill in data in tableCopy
		   for (int c = 0; c < data[t].length; c++) {
		    value = data[t][c];
		    row = tableCopy.getRow(1);
		    cell = row.getCell(c);
		 // removeCellValues(cell);
		 //   cell.setText(value);
		   }
		   document.setTable(t, tableCopy); //set tableCopy at position t instead of table

		   paragraph = document.insertNewParagraph(cursor); //insert new empty paragraph
		   cursor = setCursorToNextStartToken(paragraph.getCTP());
		  }

		  paragraph = document.insertNewParagraph(cursor);
		  run = paragraph.createRun(); 
		 
		  cursor = setCursorToNextStartToken(paragraph.getCTP());

		  FileOutputStream out = new FileOutputStream("D:/Lab/WordResult12.docx");
		  System.out.println("done");
		  document.write(out);
		  out.close();
		  document.close();
		 }
		}