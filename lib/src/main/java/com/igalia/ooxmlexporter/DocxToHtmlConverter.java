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

    private Set<DocxStyle> styleList = new HashSet<DocxStyle>();
    private XWPFDocument docx;

    public DocxToHtmlConverter(String inputFile) {
        this.inputFile = inputFile;
    }

    @Override
    public void convert(String outputFile) {
        this.outputFile = outputFile + "." + getDefaultExtension();
        try {
            docx = new XWPFDocument(new FileInputStream(inputFile));
            StringBuilder html = new StringBuilder();
            Set<String> usedStyles = new HashSet<String>();

            for (IBodyElement bodyElement : docx.getBodyElements()) {
                if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                    XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                    html.append("<p class=\"")
                            .append(paragraph.getStyle())
                            .append("\" style=\"")
                            .append(generateCSSForParagraph(paragraph))
                            .append("\">");
                    for (XWPFRun textRegion : paragraph.getRuns()) {
                        html.append("<span style='")
                                .append(generateCSSForTextRegion(textRegion))
                                .append("'>")
                                .append(textRegion.text())
                                .append("</span>");
                        usedStyles.add(textRegion.getStyle());

                        for (XWPFPicture picture : textRegion.getEmbeddedPictures()) {
                            XWPFPictureData pictureData = picture.getPictureData();
                            byte[] rawData = pictureData.getData();
                            byte[] encodedData = Base64.getEncoder().encode(rawData);
                            double pictureWidth = picture.getWidth() * 4/3; // convert from pt to px
                            html.append("<img src=\"data:")
                                    .append(pictureData.getPictureTypeEnum().getContentType())
                                    .append(";base64,")
                                    .append(new String(encodedData))
                                    .append("\" width=\"")
                                    .append(pictureWidth)
                                    .append("\" align=\"top\" style=\"float:left;\">");
                        }
                    }
                    usedStyles.add(paragraph.getStyle());
                    html.append("</p>");
                }
                if (bodyElement.getElementType() == BodyElementType.TABLE) {
                    XWPFTable table = (XWPFTable) bodyElement;
                    html.append("<table>");
                    for (XWPFTableRow row: table.getRows()) {
                        html.append("<tr>");
                        for (XWPFTableCell cell: row.getTableCells()) {
                            html.append("<td>")
                                    .append(cell.getText())
                                    .append("</td>");
                        }
                        html.append("</tr>");
                    }
                    html.append("</table>");
                }
            }

            generateStyleObjects(usedStyles);
            html.append("<style>")
                    .append(generateCSSForStyles(styleList))
                    .append("</style>");

            FileOutputStream outputHtml = new FileOutputStream(this.outputFile);
            outputHtml.write(html.toString().getBytes());
            outputHtml.close();
            docx.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private static String generateCSSForStyles(Set<DocxStyle> styleList) {
        StringBuilder css = new StringBuilder();
        for(DocxStyle style : styleList) {
            css.append(style.toCSS());
        }
        return css.toString();
    }

    private void generateStyleObjects(Set<String> usedStyles) {
        for (String styleId : usedStyles) {
            if (!styleId.isEmpty()) {
                createAndAddStyleObject(styleId);
            }
        }
    }

    private DocxStyle createAndAddStyleObject(String styleId) {
        // Return the style from the list if it was already created
        for (DocxStyle style : styleList) {
            if (style.getName().equals(styleId)) {
                return style;
            }
        }

        CTStyle ctStyle = docx.getStyles().getStyle(styleId).getCTStyle();
        DocxStyle style;

        CTString basedOn = ctStyle.getBasedOn();
        if (basedOn != null) {
            DocxStyle baseStyle = createAndAddStyleObject(basedOn.getVal());
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
            BigInteger size = (BigInteger) ctStyle.getRPr().getSzArray()[0].getVal();
            // Size in OOXML is expressed as "half points", so we divide by two.
            style.getFontFormat().setFontSize(size.doubleValue() / 2);
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
        if (ctStyle.getPPr().getJc() != null) {
            switch(ctStyle.getPPr().getJc().getVal().toString()) {
                case "start":
                    style.getFontFormat().setParagraphAlignment(ParagraphAlignment.START);
                    break;
                case "end":
                    style.getFontFormat().setParagraphAlignment(ParagraphAlignment.END);
                    break;
                case "center":
                    style.getFontFormat().setParagraphAlignment(ParagraphAlignment.CENTER);
                    break;
                case "both":
                    style.getFontFormat().setParagraphAlignment(ParagraphAlignment.BOTH);
                    break;
            }
        }

        styleList.add(style);
        return style;
    }

    private static String generateCSSForTextRegion(XWPFRun textRegion) {
        FontFormat fontFormat = new FontFormat(textRegion.getFontName(),
                textRegion.getColor(),
                textRegion.getFontSizeAsDouble(),
                textRegion.isBold(), textRegion.isItalic());
        return fontFormat.toCSS();
    }

    private static String generateCSSForParagraph(XWPFParagraph paragraph) {
        // getAlignment() falls back to LEFT if there is no value, so we need to verify
        // if the parameter is set by checking the XML directly before we call it.
        if (paragraph.getCTP().getPPr().isSetJc()) {
            FontFormat fontFormat = new FontFormat();
            fontFormat.setParagraphAlignment(paragraph.getAlignment());
            return fontFormat.toCSS();
        }
        return new String();
    }

    @Override
    public String getDefaultExtension() {
        return "html";
    }
}
