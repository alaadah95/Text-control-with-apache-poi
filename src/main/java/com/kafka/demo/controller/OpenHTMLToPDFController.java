package com.kafka.demo.controller;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.kafka.demo.controller.EditWord.Border;

@RestController
public class OpenHTMLToPDFController {

	@PostMapping("/create-bookmarks")
	public void createbookmarks() throws Exception {

		XWPFDocument document = new XWPFDocument(new FileInputStream("hello.docx"));

		List<XWPFParagraph> para = document.getParagraphs();

		/***
		 * top of page
		 */

		XmlCursor cursor = para.get(0).getCTP().newCursor();
		XWPFParagraph paragraph = para.get(0).getBody().insertNewParagraph(cursor);

		CTBookmark bookmark = paragraph.getCTP().addNewBookmarkStart();
		bookmark.setName("bookmark_3");
		bookmark.setId(BigInteger.valueOf(2));
		paragraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(2));

		cursor = para.get(0).getCTP().newCursor();
		paragraph = para.get(0).getBody().insertNewParagraph(cursor);

//		paragraph = document.createParagraph();
//		XWPFRun run = paragraph.createRun();
//		run.addBreak();
		// Bookmark the run
		bookmark = paragraph.getCTP().addNewBookmarkStart();
		bookmark.setName("bookmark_2");
		bookmark.setId(BigInteger.valueOf(1));
		paragraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(1));

		cursor = para.get(0).getCTP().newCursor();
		paragraph = para.get(0).getBody().insertNewParagraph(cursor);

//		paragraph = document.createParagraph();
//		run = paragraph.createRun();
//		run.addBreak();

		// Bookmark after the run
		bookmark = paragraph.getCTP().addNewBookmarkStart();
		bookmark.setName("bookmark_1");
		bookmark.setId(BigInteger.valueOf(0));
		paragraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(0));

		/***
		 * bottom of page
		 */

		XWPFParagraph bottomParagraph = document.createParagraph();
		CTBookmark bottomBookmark = bottomParagraph.getCTP().addNewBookmarkStart();
		bottomBookmark.setName("bookmark_4");
		bottomBookmark.setId(BigInteger.valueOf(4));
		bottomParagraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(4));

		bottomParagraph = document.createParagraph();
//		XWPFRun bottomRun = bottomParagraph.createRun();
//		bottomRun.addBreak();
		// Bookmark the run
		bottomBookmark = bottomParagraph.getCTP().addNewBookmarkStart();
		bottomBookmark.setName("bookmark_5");
		bottomBookmark.setId(BigInteger.valueOf(5));
		bottomParagraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(5));

		bottomParagraph = document.createParagraph();
//		bottomRun = bottomParagraph.createRun();
//		bottomRun.addBreak();

		// Bookmark after the run
		bottomBookmark = bottomParagraph.getCTP().addNewBookmarkStart();
		bottomBookmark.setName("bookmark_6");
		bottomBookmark.setId(BigInteger.valueOf(6));
		bottomParagraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(6));

		FileOutputStream out = new FileOutputStream("hello.docx");
		document.write(out);
		out.close();
		document.close();
	}

	@PostMapping("/replace-bookmarks")
	public void replace(@RequestParam Integer position) throws Exception {

		XWPFDocument document = new XWPFDocument(new FileInputStream("hello.docx"));
		List<XWPFParagraph> para = document.getParagraphs();
		XmlCursor cursor1 = null;
		XmlCursor cursor2 = null;
		XmlCursor cursor3 = null;
		XmlCursor cursor4 = null;
		XmlCursor cursor5 = null;
		XmlCursor cursor6 = null;
		for (XWPFParagraph paragraph : para) {
			// Here you have your paragraph;
			CTP ctp = paragraph.getCTP();
			// Get all Bookmarks and loop through them .
			List<CTBookmark> bookmarks = ctp.getBookmarkStartList();
			for (CTBookmark bookmark : bookmarks) {
				if (bookmark.getName().equals("bookmark_1")) {

					cursor1 = paragraph.getCTP().newCursor();

				} else if (bookmark.getName().equals("bookmark_2")) {
					cursor2 = paragraph.getCTP().newCursor();
				}

				else if (bookmark.getName().equals("bookmark_3")) {
					cursor3 = paragraph.getCTP().newCursor();
				} else if (bookmark.getName().equals("bookmark_4")) {
					cursor4 = paragraph.getCTP().newCursor();
				} else if (bookmark.getName().equals("bookmark_5")) {
					cursor5 = paragraph.getCTP().newCursor();
				} else if (bookmark.getName().equals("bookmark_6")) {
					cursor6 = paragraph.getCTP().newCursor();
				}
			}
		}

		if (position == 1) {
			XWPFParagraph pppp = document.insertNewParagraph(cursor1);
			XWPFRun run = pppp.createRun();
			run.setText("testing string 1-1");
		} else if (position == 2) {
			XWPFParagraph pppp = document.insertNewParagraph(cursor2);
			XWPFRun run = pppp.createRun();
			run.setText("testing string 2-2");
//			createTable(document, pppp);
		} else if (position == 3) {
			XWPFParagraph pppp = document.insertNewParagraph(cursor3);
			XWPFRun run = pppp.createRun();
			run.setText("testing string 3-3");
		} else if (position == 4) {
			XWPFParagraph pppp = document.insertNewParagraph(cursor4);
			XWPFRun run = pppp.createRun();
			run.setText("testing string 4-4");
		} else if (position == 5) {
			XWPFParagraph pppp = document.insertNewParagraph(cursor5);
			XWPFRun run = pppp.createRun();
			run.setText("testing string 5-5");
		} else if (position == 6) {
			XWPFParagraph pppp = document.insertNewParagraph(cursor6);
			XWPFRun run = pppp.createRun();
			run.setText("testing string 6-6");
		}

		FileOutputStream out = new FileOutputStream("hello.docx");
		document.write(out);
		out.close();
		document.close();
	}

	void createTable(XWPFDocument document, XWPFParagraph para) {
		EditWord classPbj = new EditWord();
		// create table
		XmlCursor cursor = para.getCTP().newCursor();
		XWPFTable table = para.getBody().insertNewTbl(cursor);
		classPbj.setTableAlign(table, ParagraphAlignment.CENTER);

		// create first row
		XWPFTableRow tableRowOne = table.getRow(0);
		XWPFTableCell column1 = tableRowOne.getCell(0);

		/// column width
		CTTblWidth cellWidth = column1.getCTTc().addNewTcPr().addNewTcW();
		CTTcPr pr = column1.getCTTc().addNewTcPr();
		pr.addNewNoWrap();
		cellWidth.setW(BigInteger.valueOf(5000));

		XWPFParagraph p = column1.addParagraph();
		p.setAlignment(ParagraphAlignment.LEFT);
		XWPFRun run = p.createRun();
		run.addBreak();
		run.setBold(true);
		run.setFontSize(16);
		run.setText("Sherif signature");

		run = p.createRun();
		run.addBreak();
		run.setBold(true);
		run.setFontSize(16);
		run.setText(" ${name} ");
		CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
		cTShd.setVal(STShd.CLEAR);
		cTShd.setColor("auto");
		cTShd.setFill("FFFFFF");

		XWPFTableCell column2 = tableRowOne.addNewTableCell();

		/// column width
		cellWidth = column2.getCTTc().addNewTcPr().addNewTcW();
		pr = column2.getCTTc().addNewTcPr();
		pr.addNewNoWrap();
		cellWidth.setW(BigInteger.valueOf(5000));

		p = column2.addParagraph();
		p.setAlignment(ParagraphAlignment.RIGHT);
		run = p.createRun();
		run.addBreak();
		run.setBold(true);
		run.setFontSize(16);
		run.setText("Alaa signature");

		run = p.createRun();
		run.addBreak();
		run.setBold(true);
		run.setFontSize(16);
		run.setText(" ${name} ");
		cTShd = run.getCTR().addNewRPr().addNewShd();
		cTShd.setVal(STShd.CLEAR);
		cTShd.setColor("auto");
		cTShd.setFill("FFFFFF");

		// remove border
		for (XWPFTableRow row : table.getRows()) {

			EditWord.setTableCellBorder(row.getCell(0), Border.TOP, STBorder.NIL);
			EditWord.setTableCellBorder(row.getCell(0), Border.BOTTOM, STBorder.NIL);
			EditWord.setTableCellBorder(row.getCell(0), Border.RIGHT, STBorder.NIL);
			EditWord.setTableCellBorder(row.getCell(0), Border.LEFT, STBorder.NIL);
			EditWord.setTableCellBorder(row.getCell(1), Border.TOP, STBorder.NIL);
			EditWord.setTableCellBorder(row.getCell(1), Border.BOTTOM, STBorder.NIL);
			EditWord.setTableCellBorder(row.getCell(1), Border.RIGHT, STBorder.NIL);
			EditWord.setTableCellBorder(row.getCell(1), Border.LEFT, STBorder.NIL);
		}

	}

}
