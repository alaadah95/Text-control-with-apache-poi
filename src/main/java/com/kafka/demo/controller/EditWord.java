package com.kafka.demo.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Base64;
import java.util.List;

import javax.annotation.PostConstruct;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class EditWord {

	@Autowired
	private FixedValuesRepository fixedRepo ;
	
	@PostMapping("/write")
	public boolean write() throws FileNotFoundException, IOException {

		String fileName = "hello.docx";
		File file = new File(fileName);
		FileInputStream input = new FileInputStream(file);

		try (XWPFDocument doc = new XWPFDocument(input)) {

//			boolean validate = doc.validateProtectionPassword("123");
//			doc.removeProtectionEnforcement();
//			System.out.println(validate);
//			XWPFParagraph p = doc.createParagraph();
//			XWPFRun run = p.createRun();

			List<XWPFParagraph> p = doc.getParagraphs();
			XWPFRun run = p.get(0).insertNewRun(0); // first paragraph, 0 is the position

			run.addBreak();
			run.setBold(true);
			run.setFontSize(30);
			run.setText("Hello");

			run = p.get(0).insertNewRun(1);
			run.setBold(true);
			run.setFontSize(30);
			run.setText(" ${name} ");
			CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
			cTShd.setVal(STShd.CLEAR);
			cTShd.setColor("auto");
			cTShd.setFill("FFFFFF");

			run = p.get(0).insertNewRun(2);
			run.setText("from the other side ");
			run.setBold(true);
			run.setFontSize(30);

			// next page
//			XWPFParagraph p2 = doc.createParagraph();
//			p2.setWordWrapped(true);
//			p2.setPageBreak(true); // new page break
//
//			XWPFRun r2 = p2.createRun();
//			r2.setFontSize(40);
//			r2.setItalic(true);
//			r2.setText("New Page");

			// save it to .docx file
			try (FileOutputStream out = new FileOutputStream(fileName)) {

				doc.write(out);
			}
			doc.close();
		}

		return true;
	}

	@PostMapping("/signature-table-bottom")
	boolean createTable() throws IOException {

		String input = "hello.docx";
		String output = "hello.docx";

		// Blank Document
		XWPFDocument document = new XWPFDocument(Files.newInputStream(Paths.get(input)));

		// create table
		XWPFTable table = document.createTable();
		setTableAlign(table, ParagraphAlignment.CENTER);

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
		run.setText("Ibrahim signature");

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

			setTableCellBorder(row.getCell(0), Border.TOP, STBorder.NIL);
			setTableCellBorder(row.getCell(0), Border.BOTTOM, STBorder.NIL);
			setTableCellBorder(row.getCell(0), Border.RIGHT, STBorder.NIL);
			setTableCellBorder(row.getCell(0), Border.LEFT, STBorder.NIL);
			setTableCellBorder(row.getCell(1), Border.TOP, STBorder.NIL);
			setTableCellBorder(row.getCell(1), Border.BOTTOM, STBorder.NIL);
			setTableCellBorder(row.getCell(1), Border.RIGHT, STBorder.NIL);
			setTableCellBorder(row.getCell(1), Border.LEFT, STBorder.NIL);
		}

		
		try (FileOutputStream out = new FileOutputStream(output)) {

			document.write(out);
			out.close();
		}
		document.close();

		return true;
	}

	@PostMapping("/signature-table-top")
	boolean createTableOntheTop(DocumentModel model) throws IOException {

		String input = "hello.docx";
		String output = "hello.docx";
		
		 
		// Blank Document
		XWPFDocument document = new XWPFDocument(Files.newInputStream(Paths.get(input)));
		// create table
		List<XWPFParagraph> para = document.getParagraphs();

		// XWPFParagraph para = document.createParagraph();
		XmlCursor cursor = para.get(0).getCTP().newCursor();

		XWPFTable table = para.get(0).getBody().insertNewTbl(cursor);
		setTableAlign(table, ParagraphAlignment.CENTER);

		// create first row
		XWPFTableRow tableRowOne = table.getRow(0);
		XWPFTableCell column1 = tableRowOne.getCell(0);

		/// column width
		CTTblWidth cellWidth = column1.getCTTc().addNewTcPr().addNewTcW();
		CTTcPr pr = column1.getCTTc().addNewTcPr();
		pr.addNewNoWrap();
		cellWidth.setW(BigInteger.valueOf(2500));

		XWPFParagraph p = column1.addParagraph();
		p.setAlignment(ParagraphAlignment.RIGHT);
		XWPFRun run = p.createRun();
		
		// set bidirectional text support on (arabic)
//		  CTP ctp = p.getCTP();
//		  CTPPr ctppr = ctp.getPPr();
//		  if (ctppr == null) ctppr = ctp.addNewPPr();
//		  ctppr.addNewBidi().setVal(STOnOff.ON);
		
//		run.addBreak();
		run.setBold(false);
		run.setFontSize(16);
		run.setText(model.getDate());
		 
//		run = p.createRun();
//		run.addBreak();
//		run.setBold(true);
//		run.setFontSize(16);
//		run.setText(" ${name} ");
//		CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
//		cTShd.setVal(STShd.CLEAR);
//		cTShd.setColor("auto");
//		cTShd.setFill("FFFFFF");
		
		
		XWPFTableCell column2 = tableRowOne.addNewTableCell();  
		/// column width
		cellWidth = column2.getCTTc().addNewTcPr().addNewTcW();
		pr = column2.getCTTc().addNewTcPr();
		pr.addNewNoWrap();
		cellWidth.setW(BigInteger.valueOf(3500));
		
		p = column2.addParagraph();
		p.setAlignment(ParagraphAlignment.LEFT);
		run = p.createRun();
		run.setBold(false);
		run.setFontSize(16);
		run.setText(fixedRepo.findById(1).get().getName());
		
		
		XWPFTableCell column3 = tableRowOne.addNewTableCell();
		/// column width
		cellWidth = column3.getCTTc().addNewTcPr().addNewTcW();
		pr = column3.getCTTc().addNewTcPr();
		pr.addNewNoWrap();
		cellWidth.setW(BigInteger.valueOf(4000));

		p = column3.addParagraph();
		p.setAlignment(ParagraphAlignment.LEFT);
		// set bidirectional text support on (arabic)
		CTP ctp1 = p.getCTP();
		CTPPr ctppr2 = ctp1.getPPr();
		  if (ctppr2 == null) ctppr2 = ctp1.addNewPPr();
		  ctppr2.addNewBidi().setVal(STOnOff.ON);
		
		run = p.createRun();
		for (Integer manager : model.getManagers()) {
		
			run.setBold(false);
			run.setFontSize(12);
			run.setText(fixedRepo.findById(manager).get().getName());
			run.addBreak();
		
		}

//		run = p.createRun();
//		run.addBreak();
//		run.setBold(true);
//		run.setFontSize(16);
//		run.setText(" ${name} ");
//		cTShd = run.getCTR().addNewRPr().addNewShd();
//		cTShd.setVal(STShd.CLEAR);
//		cTShd.setColor("auto");
//		cTShd.setFill("FFFFFF");

		
		///Row 2
		table.createRow();
		XWPFTableRow tableRowtwo = table.getRow(1);
		XWPFParagraph p4 = tableRowtwo.getCell(2).getParagraphs().get(0);
		p4.setAlignment(ParagraphAlignment.LEFT);
		XWPFRun r4 = p4.createRun();
		r4.setBold(true);
		r4.setText("الموضوع");
		
		p4 = tableRowtwo.getCell(1).getParagraphs().get(0);
		p4.setAlignment(ParagraphAlignment.LEFT);
		
		// set bidirectional text support on (arabic)
		CTP ctp2 = p4.getCTP();
		CTPPr ctppr3 = ctp2.getPPr();
		  if (ctppr3 == null) ctppr3 = ctp2.addNewPPr();
		  ctppr3.addNewBidi().setVal(STOnOff.ON);
		
		r4 = p4.createRun();
		r4.setBold(true);
		r4.setText("test the world");
		
				
		// remove border
		for (XWPFTableRow row : table.getRows()) {

			setTableCellBorder(row.getCell(0), Border.TOP, STBorder.NIL);
			setTableCellBorder(row.getCell(0), Border.BOTTOM, STBorder.NIL);
			setTableCellBorder(row.getCell(0), Border.RIGHT, STBorder.NIL);
			setTableCellBorder(row.getCell(0), Border.LEFT, STBorder.NIL);
			setTableCellBorder(row.getCell(1), Border.TOP, STBorder.NIL);
			setTableCellBorder(row.getCell(1), Border.BOTTOM, STBorder.NIL);
			setTableCellBorder(row.getCell(1), Border.RIGHT, STBorder.NIL);
			setTableCellBorder(row.getCell(1), Border.LEFT, STBorder.NIL);
			setTableCellBorder(row.getCell(2), Border.TOP, STBorder.NIL);
			setTableCellBorder(row.getCell(2), Border.BOTTOM, STBorder.NIL);
			setTableCellBorder(row.getCell(2), Border.RIGHT, STBorder.NIL);
			setTableCellBorder(row.getCell(2), Border.LEFT, STBorder.NIL);
		}

		try (FileOutputStream out = new FileOutputStream(output)) {

			document.write(out);
			out.close();
		}
		document.close();

		return true;
	}

	@PostMapping("/edit")
	public boolean update() throws FileNotFoundException, IOException {

		String fileName = "hello.docx";

		updateDocument(fileName, fileName, " Alaa95 - lolo signature ");

		return true;
	}

	@PostMapping("/edit-with-image")
	public boolean updateWithImage() throws FileNotFoundException, IOException, InvalidFormatException {

		String fileName = "hello.docx";

		updateDocumentWithImage(fileName, fileName, " ");

		return true;
	}

	@PostMapping("/print-with-signature")
	public String print_with_signature_MS(@RequestBody DocumentModel model)
			throws IOException, InvalidFormatException {

		System.out.println(model.getDocument());
		byte[] decodedBytes = Base64.getDecoder().decode(model.getDocument());
		Files.write(Paths.get("hello.docx"), decodedBytes);
		createTable();
		clearHeader();
		addHeaderAndFooterImage();
		createTableOntheTop(model);
		updateWithImage();

		return encoder("hello.docx");
	}

	@PostMapping("/clear-headers")
	public boolean clearHeader() throws FileNotFoundException, IOException {

		String inFilePath = "hello.docx";
		String outFilePath = "hello.docx";

		XWPFDocument document = new XWPFDocument(new FileInputStream(inFilePath));

		for (XWPFHeader header : document.getHeaderList()) {
			header.setHeaderFooter(
					org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr.Factory.newInstance());
		}
		for (XWPFFooter footer : document.getFooterList()) {
			footer.setHeaderFooter(
					org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr.Factory.newInstance());
		}

		FileOutputStream out = new FileOutputStream(outFilePath);
		document.write(out);
		out.close();
		document.close();
		return true;
	}
	 
	
	public static String encoder(String filePath) {
		String base64File = "";
		File file = new File(filePath);
		try (FileInputStream imageInFile = new FileInputStream(file)) {
			// Reading a file from file system
			byte fileData[] = new byte[(int) file.length()];
			imageInFile.read(fileData);
			base64File = Base64.getEncoder().encodeToString(fileData);
		} catch (FileNotFoundException e) {
			System.out.println("File not found" + e);
		} catch (IOException ioe) {
			System.out.println("Exception while reading the file " + ioe);
		}
		return base64File;
	}

	public enum Border {
		LEFT, TOP, BOTTOM, RIGHT
	}

	static void setTableCellBorder(XWPFTableCell cell, Border border, STBorder.Enum type) {
		CTTc tc = cell.getCTTc();
		CTTcPr tcPr = tc.getTcPr();
		if (tcPr == null)
			tcPr = tc.addNewTcPr();
		CTTcBorders tcBorders = tcPr.getTcBorders();
		if (tcBorders == null)
			tcBorders = tcPr.addNewTcBorders();
		if (border == Border.LEFT) {
			CTBorder left = tcBorders.getLeft();
			if (left == null)
				left = tcBorders.addNewLeft();
			left.setVal(type);
		} else if (border == Border.TOP) {
			CTBorder top = tcBorders.getTop();
			if (top == null)
				top = tcBorders.addNewTop();
			top.setVal(type);
		} else if (border == Border.BOTTOM) {
			CTBorder bottom = tcBorders.getBottom();
			if (bottom == null)
				bottom = tcBorders.addNewBottom();
			bottom.setVal(type);
		} else if (border == Border.RIGHT) {
			CTBorder right = tcBorders.getRight();
			if (right == null)
				right = tcBorders.addNewRight();
			right.setVal(type);
		}
	}

	public void setTableAlign(XWPFTable table, ParagraphAlignment align) {
		CTTblPr tblPr = table.getCTTbl().getTblPr();
		CTJc jc = (tblPr.isSetJc() ? tblPr.getJc() : tblPr.addNewJc());
		jc.setVal(STJc.CENTER);
	}

	private void updateDocument(String input, String output, String name) throws IOException {

		try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(input)))) {

			List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
			// Iterate over paragraph list and check for the replaceable text in each
			// paragraph
			for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
				replaceParagraph(xwpfParagraph, "${name}", name);
			}

			for (XWPFTable table : doc.getTables()) {
				replaceTable(table, "${name}", name);
			}

			// save the docs
			try (FileOutputStream out = new FileOutputStream(output)) {
				doc.write(out);
			}
			doc.close();
		}

	}

	public XWPFDocument replacePOI(XWPFDocument doc, String placeHolder, String replaceText) {
		// REPLACE ALL HEADERS
		for (XWPFHeader header : doc.getHeaderList())
			replaceAllBodyElements(header.getBodyElements(), placeHolder, replaceText);
		// REPLACE BODY
		replaceAllBodyElements(doc.getBodyElements(), placeHolder, replaceText);
		return doc;
	}

	private void replaceAllBodyElements(List<IBodyElement> bodyElements, String placeHolder, String replaceText) {
		for (IBodyElement bodyElement : bodyElements) {
			if (bodyElement.getElementType().compareTo(BodyElementType.PARAGRAPH) == 0)
				replaceParagraph((XWPFParagraph) bodyElement, placeHolder, replaceText);
			if (bodyElement.getElementType().compareTo(BodyElementType.TABLE) == 0)
				replaceTable((XWPFTable) bodyElement, placeHolder, replaceText);
		}
	}

	private void replaceTable(XWPFTable table, String placeHolder, String replaceText) {
		for (XWPFTableRow row : table.getRows()) {
			for (XWPFTableCell cell : row.getTableCells()) {
				for (IBodyElement bodyElement : cell.getBodyElements()) {
					if (bodyElement.getElementType().compareTo(BodyElementType.PARAGRAPH) == 0) {
						replaceParagraph((XWPFParagraph) bodyElement, placeHolder, replaceText);
					}
					if (bodyElement.getElementType().compareTo(BodyElementType.TABLE) == 0) {
						replaceTable((XWPFTable) bodyElement, placeHolder, replaceText);
					}
				}
			}
		}
	}

	private void replaceParagraph(XWPFParagraph paragraph, String placeHolder, String replaceText) {
		for (XWPFRun r : paragraph.getRuns()) {
			String text = r.getText(r.getTextPosition());
			if (text != null && text.contains(placeHolder)) {
				text = text.replace(placeHolder, replaceText);
				r.setText(text, 0);
			}
		}
	}

	private void replaceTableWithImage(XWPFTable table, String placeHolder, String replaceText, String imgFile)
			throws InvalidFormatException, IOException {
		for (XWPFTableRow row : table.getRows()) {
			for (XWPFTableCell cell : row.getTableCells()) {
				for (IBodyElement bodyElement : cell.getBodyElements()) {
					if (bodyElement.getElementType().compareTo(BodyElementType.PARAGRAPH) == 0) {
						replaceParagraphWithImage((XWPFParagraph) bodyElement, placeHolder, replaceText, imgFile);
					}
					if (bodyElement.getElementType().compareTo(BodyElementType.TABLE) == 0) {
						replaceTable((XWPFTable) bodyElement, placeHolder, replaceText);
					}
				}
			}
		}
	}

	private void replaceParagraphWithImage(XWPFParagraph paragraph, String placeHolder, String replaceText,
			String imgFile) throws InvalidFormatException, IOException {
		for (XWPFRun r : paragraph.getRuns()) {
			String text = r.getText(r.getTextPosition());
			if (text != null && text.contains(placeHolder)) {
				text = text.replace(placeHolder, replaceText);
				try (FileInputStream is = new FileInputStream(imgFile)) {
					r.addPicture(is, Document.PICTURE_TYPE_PNG, // png file
							imgFile, Units.toEMU(150), Units.toEMU(50)); // 150x200 pixels
					r.addBreak();
					is.close();
				}
				r.setText(text, 0);
			}
		}
	}

	private void updateDocumentWithImage(String input, String output, String name)
			throws IOException, InvalidFormatException {

		String imgFile = "download.png";

		try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(input)))) {

			List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
			// Iterate over paragraph list and check for the replaceable text in each
			// paragraph
			for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
				replaceParagraphWithImage(xwpfParagraph, "${name}", name, imgFile);
			}

			for (XWPFTable table : doc.getTables()) {
				replaceTableWithImage(table, "${name}", name, imgFile);
			}

			// save the docs
			try (FileOutputStream out = new FileOutputStream(output)) {
				doc.write(out);
			}
			doc.close();
		}

	}

	private boolean addHeaderAndFooterImage() throws InvalidFormatException, IOException {
		String inFilePath = "hello.docx";
		String outFilePath = "hello.docx";

		XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(inFilePath)));

		// the body content
		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();

		// create header
		XWPFHeader header = doc.createHeader(HeaderFooterType.DEFAULT);

		// header's first paragraph
		paragraph = header.getParagraphArray(0);
		if (paragraph == null)
			paragraph = header.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);

		run = paragraph.createRun();

		String headerImg = "header.png";
		FileInputStream in = new FileInputStream(headerImg);
		run.addPicture(in, Document.PICTURE_TYPE_PNG, headerImg, Units.toEMU(595), Units.toEMU(80));
		in.close();

		XWPFFooter footer = doc.createFooter(HeaderFooterType.DEFAULT);

		paragraph = footer.getParagraphArray(0);
		if (paragraph == null)
			paragraph = footer.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);

		run = paragraph.createRun();

		String footerImg = "footer.png";
		in = new FileInputStream(footerImg);
		run.addPicture(in, Document.PICTURE_TYPE_PNG, footerImg, Units.toEMU(595), Units.toEMU(80));
		in.close();

		FileOutputStream out = new FileOutputStream(outFilePath);
		doc.write(out);
		doc.close();
		out.close();
		return true;
	}

	@PostConstruct
	public void EditWord() {
		FixedValues entity = new FixedValues();
		entity.setId(1);
		entity.setName(" : الموافق   ");
		fixedRepo.save(entity);
		
		entity = new FixedValues();
		entity.setId(2);
		entity.setName("المدراء التنفيذين");
		fixedRepo.save(entity);
		
		entity = new FixedValues();
		entity.setId(3);
		entity.setName("علاء محمود ضاهر");
		fixedRepo.save(entity);
		
		entity = new FixedValues();
		entity.setId(4);
		entity.setName("كريم رشدي محمد");
		fixedRepo.save(entity);
	}
	
	
	
}
