package com.igalia.ooxmlexporter;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.HashSet;
import java.util.Set;

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
            Set<String> usedStyles = new HashSet<String>();
            Set<DocxStyle> styles = new HashSet<DocxStyle>();

            for (IBodyElement bodyElement : docx.getBodyElements()) {
                if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                    html.append("<p>");
                    for (XWPFRun textRegion : ((XWPFParagraph) bodyElement).getRuns()) {
                        html.append("<span style='")
                                .append(generateCSSForTextRegion(textRegion))
                                .append("'>")
                                .append(textRegion.text())
                                .append("</span>");
                        usedStyles.add(textRegion.getStyle());
                    }
                    usedStyles.add(((XWPFParagraph) bodyElement).getStyle());
                    html.append("</p>");
                }
            }
            expandStyleList(docx.getStyles(), usedStyles, styles);
            System.out.println(generateCSSForStyles(docx.getStyles(), styles));
            html.append("<style>")
                    .append(generateCSSForStyles(docx.getStyles(), styles))
                    .append("</style>");

            FileOutputStream outputHtml = new FileOutputStream(outputFile);
            outputHtml.write(html.toString().getBytes());
            outputHtml.close();
            docx.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private String generateCSSForStyles(XWPFStyles styles, Set<DocxStyle> styleList) {
        StringBuilder css = new StringBuilder();
        for(DocxStyle style : styleList) {
            css.append(style.toCSS());
        }
        return css.toString();
    }

    private void expandStyleList(XWPFStyles styles, Set<String> usedStyles, Set<DocxStyle> styleList) {
        Set<String> originalList = new HashSet<String>(usedStyles);
        for (String styleId : originalList) {
            if (!styleId.isEmpty()) {
                styleList.add(new DocxStyle(styleId));
                addBasedOnStyles(styleId, styles, styleList);
            }
        }
    }

    private void addBasedOnStyles(String styleId, XWPFStyles styles, Set<DocxStyle> styleList) {
        CTStyle ctStyle = styles.getStyle(styleId).getCTStyle();
        DocxStyle style = new DocxStyle(styleId);
        if (ctStyle.getRPr().getRFontsArray().length > 0) {
            style.setFontName(ctStyle.getRPr().getRFontsArray()[0].getAscii());
            style.setFontSize((BigInteger) ctStyle.getRPr().getSzArray()[0].getVal());
        }
        styleList.add(style);

        CTString basedOn = ctStyle.getBasedOn();
        if (basedOn != null) {
            addBasedOnStyles(basedOn.getVal(), styles, styleList);
        }
    }

    public String generateCSSForTextRegion(XWPFRun textRegion) {
        StringBuilder css = new StringBuilder();
        String fontName = textRegion.getFontName();
        if (fontName != null) {
            css.append("font-family: \"").append(fontName).append("\"; ");
        }
        if (textRegion.isBold()) {
            css.append("font-weight: bold; ");
        }
        if (textRegion.isItalic()) {
            css.append("font-style: italic; ");
        }
        Double fontSize = textRegion.getFontSizeAsDouble();
        if (fontSize != null) {
            css.append("font-size: ").append(fontSize).append("px; ");
        }
        String fontColor = textRegion.getColor();
        if (fontColor != null) {
            css.append("color: #").append(fontColor).append("; ");
        }
        return css.toString();
    }

    @Override
    public String getDefaultExtension() {
        return "html";
    }
}
