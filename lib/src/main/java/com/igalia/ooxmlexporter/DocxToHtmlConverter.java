package com.igalia.ooxmlexporter;

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
                    html.append("<p>");
                    for (XWPFRun textRegion : ((XWPFParagraph) bodyElement).getRuns()) {
                        html.append("<span style='")
                                .append(generateCSSForTextRegion(textRegion))
                                .append("'>")
                                .append(textRegion.text())
                                .append("</span>");
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

    public String generateCSSForTextRegion(XWPFRun textRegion) {
        StringBuilder css = new StringBuilder();
        String fontName = textRegion.getFontName();
        if (fontName != null) {
            css.append("font-family: \"" + fontName + "\"; ");
        }
        if (textRegion.isBold()) {
            css.append("font-weight: bold; ");
        }
        if (textRegion.isItalic()) {
            css.append("font-style: italic; ");
        }
        Double fontSize = textRegion.getFontSizeAsDouble();
        if (fontSize != null) {
            css.append("font-size: " + fontSize + "px; ");
        }
        String fontColor = textRegion.getColor();
        if (fontColor != null) {
            css.append("color: #" + fontColor + "; ");
        }
        return css.toString();
    }

    @Override
    public String getDefaultExtension() {
        return "html";
    }
}
