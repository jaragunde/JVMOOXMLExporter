package com.igalia.ooxmlexporter;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.Base64;
import java.util.HashSet;
import java.util.Set;

public class DocxToHtmlConverter implements DocumentConverter {
    private String inputFile;
    private String outputFile;

    public DocxToHtmlConverter(String inputFile) {
        this.inputFile = inputFile;
    }

    @Override
    public void convert(String outputFile) {
        this.outputFile = outputFile + "." + getDefaultExtension();
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

            for (XWPFPictureData pictureData : docx.getAllPictures()) {
                byte[] rawData = pictureData.getData();
                byte[] encodedData = Base64.getEncoder().encode(rawData);
                html.append("<img src=\"data:")
                        .append(pictureData.getPictureTypeEnum().getContentType())
                        .append(";base64,")
                        .append(new String(encodedData))
                        .append("\">");
            }

            expandStyleList(docx.getStyles(), usedStyles, styles);
            System.out.println(generateCSSForStyles(docx.getStyles(), styles));
            html.append("<style>")
                    .append(generateCSSForStyles(docx.getStyles(), styles))
                    .append("</style>");

            FileOutputStream outputHtml = new FileOutputStream(this.outputFile);
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
            style.getFontFormat().setFontName(ctStyle.getRPr().getRFontsArray()[0].getAscii());
        }
        if (ctStyle.getRPr().getSzArray().length > 0) {
            style.getFontFormat().setFontSize((BigInteger) ctStyle.getRPr().getSzArray()[0].getVal());
        }
        if (ctStyle.getRPr().getBArray().length > 0) {
            // The presence of a `<w:b/>` tag indicates that bold is set.
            style.getFontFormat().setBold(true);
        }
        if (ctStyle.getRPr().getIArray().length > 0) {
            // The presence of a `<w:b/>` tag indicates that bold is set.
            style.getFontFormat().setItalic(true);
        }
        if (ctStyle.getRPr().getColorArray().length > 0) {
            // xgetVal() returns a STHexColorImpl object, which can output a string with RRGGBB values.
            style.getFontFormat().setFontColor(ctStyle.getRPr().getColorArray()[0].xgetVal().getStringValue());
        }
        styleList.add(style);
        return style;
    }

    public String generateCSSForTextRegion(XWPFRun textRegion) {
        Double fontSize = textRegion.getFontSizeAsDouble();
        FontFormat fontFormat = new FontFormat(textRegion.getFontName(),
                textRegion.getColor(),
                fontSize != null ? BigInteger.valueOf(fontSize.longValue()) : null,
                textRegion.isBold(), textRegion.isItalic());
        return fontFormat.toCSS();
    }

    @Override
    public String getDefaultExtension() {
        return "html";
    }
}
