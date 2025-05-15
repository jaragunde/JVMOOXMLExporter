package com.igalia.ooxmlexporter;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.xwpf.usermodel.*;

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
                    html.append("<p>");
                    for (XWPFRun textRegion : ((XWPFParagraph) bodyElement).getRuns()) {
                        html.append("<span style='");
                        String fontName = textRegion.getFontName();
                        if (fontName != null) {
                            html.append("font-family: \"" + fontName + "\"; ");
                        }
                        Double fontSize = textRegion.getFontSizeAsDouble();
                        if (fontSize != null) {
                            html.append("font-size: " + fontSize + "px; ");
                        }
                        html.append("'>").append(textRegion.text()).append("</span>");
                    }
                    html.append("</p>");
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
