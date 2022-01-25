package zama.learning;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Calendar;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.text.BadElementException;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImage;
import com.itextpdf.text.pdf.PdfIndirectObject;
import com.itextpdf.text.pdf.PdfName;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;

public class PDFConverterDemo {
	private static final Logger log = LoggerFactory.getLogger(PDFConverterDemo.class.getName());
	private static Long uniqueIdentifier = Calendar.getInstance().getTimeInMillis();
	private static String DOC_HOME = "C:\\Users\\91988\\Desktop\\testing\\pdf-conversion\\";
	private static String HTML_FILE_NAME = DOC_HOME+"sample.html"; 
	private static String DOC_FILE_NAME = DOC_HOME+"1-2022-900271.doc"; 
	private static String DOCX_FILE_NAME = DOC_HOME+"nda-template.docx";
	private static String DOCX_FILE_NAME_SPACES = DOC_HOME+"nda_updated.docx";
	private static String DOCX_FILE_NAME2 = DOC_HOME+"nda-template_openxml.docx";
	private static String PDF_TEMPLATE = DOC_HOME+"nda-template_doc.pdf"; 
	private static String PDF_FILE_NAME1 = DOC_HOME+"htmlconvertor1_7"+uniqueIdentifier+"pdf"; 
	private static String PDF_FILE_NAME2 = DOC_HOME+"docxconvertor1_7"+uniqueIdentifier+".pdf"; 
	private static String PDF_FILE_NAME3 = DOC_HOME+"docxconvertor_poi_"+uniqueIdentifier+".pdf"; 
	private static String PDF_FILE_NAME4 = DOC_HOME+"docxconvertor_poi_final.pdf"; 
	private static String LOGO = "https://csscorpuat.bobeprocure.com/upeg/b4img/csscorp-logo.png";

	public static void main(String[] args) throws Exception {
		//Docx to PDF
		convertDocxToPDFUsingPoi(DOCX_FILE_NAME, PDF_FILE_NAME3);

		//Convert .docx --> HTML --> PDF (Content is shifting to right) 
		//writeToPDFUsingHtmlConvertor(convertDocxToHTML(DOCX_FILE_NAME));
	}

	private static String convertDocxToHTML(String docPath) throws FileNotFoundException, IOException {
		InputStream in= new FileInputStream(new File(docPath));
		XWPFDocument document = new XWPFDocument(in);
		XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(new File("DOCX")));
		OutputStream out = new ByteArrayOutputStream();
		XHTMLConverter.getInstance().convert(document, out, options);
		String html=out.toString();
		log.info(html);
		return html;
	}

	//This is working
	private static void writeToPDFUsingHtmlConvertor(String template) throws IOException, FileNotFoundException {
		ConverterProperties props = new ConverterProperties();
		//props.setCssApplierFactory(null)
		HtmlConverter.convertToPdf(template, new FileOutputStream(PDF_FILE_NAME2), props);
	}

	//This is not working
	private static void writeToPDFUsingXMLConvertor(String template)
			throws DocumentException, FileNotFoundException, IOException {
		Document document = new Document();
		PdfWriter writer;

		String dst = "C:\\Users\\91988\\Downloads\\";
		java.io.File folder = new java.io.File(dst);
		if (!folder.exists()) {
			folder.mkdirs();
		}
		String contractId = "1_7";
		java.io.File file = new java.io.File(dst, contractId  + ".pdf");

		writer = PdfWriter.getInstance(document, new FileOutputStream(file));
		writer.setPdfVersion(PdfWriter.VERSION_1_7);
		document.open();
		XMLWorkerHelper.getInstance().parseXHtml(writer, document, new ByteArrayInputStream(template.getBytes()));
		document.close();
	}

	private static String readFile(String filePath) throws IOException {
		File file = new File(filePath);
		StringBuilder fileContents = new StringBuilder((int)file.length());        

		try (Scanner scanner = new Scanner(file)) {
			while(scanner.hasNextLine()) {
				fileContents.append(scanner.nextLine() + System.lineSeparator());
			}
			return fileContents.toString();
		}
	}

	public static String addExtraLinesToDocx(String fileName) {
		log.info("Reading: {}", fileName);

		String updatedFilePath = "";
		try {
			File file = new File(fileName);
			FileInputStream fis = new FileInputStream(file.getAbsolutePath());

			XWPFDocument document = new XWPFDocument(fis);

			List<XWPFParagraph> paragraphs = document.getParagraphs();

			XWPFRun run = paragraphs.get(0).createRun();
			for(int i=0; i<4; i++) {
				run.addCarriageReturn();
				run.addBreak();
			}

			updatedFilePath = file.getParent()+"\\"+Calendar.getInstance().getTimeInMillis()+"_updated.docx";
			FileOutputStream out = new FileOutputStream(new File(updatedFilePath));

			for (XWPFParagraph para : paragraphs) { 
				log.info(para.getText()); 
			}

			document.write(out);
			out.close();
			log.info("Successfully updated doc: {}", updatedFilePath);

			fis.close();
		} catch (Exception e) {
			log.error("Error while adding extra lines: {}", e);
		}
		return updatedFilePath;
	}


	public static void convertDocxToPDFUsingPoi(String docPath, String pdfPath) {
		try {
			//Add extra lines to add logo
			log.info("Adding extra lines for logo in docx");
			String updatedFilePath = addExtraLinesToDocx(docPath);

			log.info("Converting docx to PDF");
			InputStream doc = new FileInputStream(new File(updatedFilePath));
			XWPFDocument document = new XWPFDocument(doc);

			PdfOptions options = PdfOptions.create();
			OutputStream out = new FileOutputStream(new File(pdfPath));
			PdfConverter.getInstance().convert(document, out, options);
			out.close();

			log.info("Adding logo to updated PDF");
			manipulatePdf(pdfPath, PDF_FILE_NAME4);
		} catch (Exception ex) {
			log.error("{}", ex);
		}
	}

	private static void addLogoToPDF(String docPath)
			throws BadElementException, MalformedURLException, IOException, DocumentException, FileNotFoundException {
		Image image = Image.getInstance(new URL("https://csscorpuat.bobeprocure.com/upeg/b4img/csscorp-logo.png"));

		Document newdocument = new Document();
		PdfWriter.getInstance(newdocument, new FileOutputStream(docPath));
		newdocument.open();
		newdocument.add(image);
		newdocument.close();
		log.info("Written image to pdf");
	}

	private static void manipulatePdf(String src, String dest) throws IOException, DocumentException {
		PdfReader reader = new PdfReader(src);
		PdfStamper stamper = new PdfStamper(reader, new FileOutputStream(dest));
		Image image = Image.getInstance(LOGO);
		PdfImage stream = new PdfImage(image, "", null);
		stream.put(new PdfName("NDA"), new PdfName("NDA"));
		PdfIndirectObject ref = stamper.getWriter().addToBody(stream);
		image.setDirectReference(ref.getIndirectReference());
		image.setAbsolutePosition(0, 700);
		image.setWidthPercentage(0.4f);
		PdfContentByte over = stamper.getOverContent(1);
		over.addImage(image);
		stamper.close();
		reader.close();
		log.info("Updated PDF");
	}
}
