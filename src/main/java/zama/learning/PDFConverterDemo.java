package zama.learning;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;

public class PDFConverterDemo {
	private static final Logger log = LoggerFactory.getLogger(PDFConverterDemo.class.getName());

	public static void main(String[] args) throws Exception {
		String template = readFile();
		log.info("template: {}", template);
		
		writeToPDFUsingHtmlConvertor(template);
		//writeToPDFUsingXMLConvertor(template);
	}
	
	//This is working
	private static void writeToPDFUsingHtmlConvertor(String template) throws IOException, FileNotFoundException {
		HtmlConverter.convertToPdf(template, new FileOutputStream("C:\\Users\\91988\\Downloads\\htmlconvertor1_7.pdf"));
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
	
	private static String readFile() throws IOException {
	    File file = new File("C:\\Users\\91988\\Downloads\\sample.html");
	    StringBuilder fileContents = new StringBuilder((int)file.length());        

	    try (Scanner scanner = new Scanner(file)) {
	        while(scanner.hasNextLine()) {
	            fileContents.append(scanner.nextLine() + System.lineSeparator());
	        }
	        return fileContents.toString();
	    }
	}
}
