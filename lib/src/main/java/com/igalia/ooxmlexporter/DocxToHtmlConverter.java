package com.igalia.ooxmlexporter;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class DocxToHtmlConverter implements DocumentConverter {
    private String inputFile;
    private String outputFile;

    public DocxToHtmlConverter(String inputFile, String outputFile) {
        this.inputFile = inputFile;
        this.outputFile = outputFile + "." + getDefaultExtension();
    }

    @Override
    public void convert() {
        try {
            XWPFDocument docx = new XWPFDocument(new FileInputStream(inputFile));
            StringBuilder html = new StringBuilder();

            for (IBodyElement bodyElement : docx.getBodyElements()) {
                if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                    XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                    html.append("<p>").append(paragraph.getText()).append("</p>");
                }
            }

            FileOutputStream outputHtml = new FileOutputStream(outputFile);
            outputHtml.write(html.toString().getBytes());
            outputHtml.close();
            docx.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    @Override
    public String getDefaultExtension() {
        return "html";
    }
}
