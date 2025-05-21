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
                    XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                    html.append("<p class=\"").append(paragraph.getStyle()).append("\">");
                    for (XWPFRun textRegion : paragraph.getRuns()) {
                        html.append("<span style='")
                                .append(generateCSSForTextRegion(textRegion))
                                .append("'>")
                                .append(textRegion.text())
                                .append("</span>");
                        usedStyles.add(textRegion.getStyle());
                    }
                    usedStyles.add(paragraph.getStyle());
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
                createStyleObjectAndBasedOn(styleId, styles, styleList);
            }
        }
    }

    private DocxStyle createStyleObjectAndBasedOn(String styleId, XWPFStyles styles, Set<DocxStyle> styleList) {
        // Return the style from the list if it was already created
        for (DocxStyle style : styleList) {
            if (style.getName().equals(styleId)) {
                return style;
            }
        }

        CTStyle ctStyle = styles.getStyle(styleId).getCTStyle();
        DocxStyle style;

        CTString basedOn = ctStyle.getBasedOn();
        if (basedOn != null) {
            DocxStyle baseStyle = createStyleObjectAndBasedOn(basedOn.getVal(), styles, styleList);
            style = new DocxStyle(styleId, baseStyle);
        }
        else {
            style = new DocxStyle(styleId);
        }
        // A style can override attributes from the base style, that's why we do this
        // in second place.
        if (ctStyle.getRPr().getRFontsArray().length > 0) {
            style.setFontName(ctStyle.getRPr().getRFontsArray()[0].getAscii());
            style.setFontSize((BigInteger) ctStyle.getRPr().getSzArray()[0].getVal());
        }
        if (ctStyle.getRPr().getBArray().length > 0) {
            // The presence of a `<w:b/>` tag indicates that bold is set.
            style.setBold(true);
        }
        styleList.add(style);
        return style;
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
