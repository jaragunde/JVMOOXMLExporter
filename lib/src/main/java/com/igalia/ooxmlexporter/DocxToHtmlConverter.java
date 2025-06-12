package com.igalia.ooxmlexporter;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.Base64;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class DocxToHtmlConverter implements DocumentConverter {
    private String inputFile;
    private String outputFile;

    private Set<DocxStyle> styleList = new HashSet<DocxStyle>();
    private XWPFDocument docx;

    // 0 means top-level list, 1 is first nested level, etc.
    // -1 means the list is closed.
    private int currentListLevel;

    public DocxToHtmlConverter(String inputFile) {
        this.inputFile = inputFile;
    }

    @Override
    public void convert(String outputFile) {
        this.outputFile = outputFile + "." + getDefaultExtension();
        try {
            docx = new XWPFDocument(new FileInputStream(inputFile));
            StringBuilder html = new StringBuilder();

            CTSectPr sectionProperties = docx.getDocument().getBody().getSectPr();
            int pageWidthVal = ((BigInteger) sectionProperties.getPgSz().getW()).intValue();
            int leftMarginVal = ((BigInteger) sectionProperties.getPgMar().getLeft()).intValue();
            int rightMarginVal = ((BigInteger) sectionProperties.getPgMar().getRight()).intValue();
            // pageWidthVal is the page size, but we need the actual width of the content
            int contentWidth = pageWidthVal - leftMarginVal - rightMarginVal;
            // convert from twentieths of a point to pixels
            double width = contentWidth / 20 * 4 / 3 ;
            double paddingLeft = leftMarginVal / 20 * 4 / 3 ;
            double paddingRight = rightMarginVal / 20 * 4 / 3 ;
            html.append("<div style='")
                    .append("width: ").append(width).append("px; ")
                    .append("padding-left: ").append(paddingLeft).append("px; ")
                    .append("padding-right: ").append(paddingRight).append("px;'>");

            for (IBodyElement bodyElement : docx.getBodyElements()) {
                if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                    html.append(generateHTMLForParagraph((XWPFParagraph) bodyElement));
                }
                if (bodyElement.getElementType() == BodyElementType.TABLE) {
                    XWPFTable table = (XWPFTable) bodyElement;
                    html.append("<table style='")
                            .append(generateCSSForTable(table))
                            .append("'>");
                    for (XWPFTableRow row: table.getRows()) {
                        html.append("<tr>");
                        for (XWPFTableCell cell: row.getTableCells()) {
                            html.append("<td style='")
                                    .append(generateCSSForTableCell(cell))
                                    .append("'>");
                            for (XWPFParagraph paragraph : cell.getParagraphs()) {
                                html.append(generateHTMLForParagraph(paragraph));
                            }
                            html.append("</td>");
                        }
                        html.append("</tr>");
                    }
                    html.append("</table>");
                }
            }

            generateStyleObjects(docx.getStyles().getStyles());
            html.append("<style>")
                    .append(generateCSSForStyles(styleList))
                    .append("</style>");

            html.append("</div>");

            FileOutputStream outputHtml = new FileOutputStream(this.outputFile);
            outputHtml.write(html.toString().getBytes());
            outputHtml.close();
            docx.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private String generateHTMLForParagraph(XWPFParagraph paragraph) {
        StringBuilder html = new StringBuilder();

        // Paragraphs with style Heading 1, 2, etc. have an assigned NumIlvl, although their NumFmt is "none".
        // In that special case, we force newListLevel to be -1 (list closed).
        boolean isParagraphInsideList = (paragraph.getNumFmt() != null)
                && !paragraph.getNumFmt().equals("none");
        int newListLevel = isParagraphInsideList ? paragraph.getNumIlvl().intValue() : -1;
        while (currentListLevel > newListLevel) {
            html.append("</ul>");
            currentListLevel--;
        }
        while (currentListLevel < newListLevel) {
            html.append("<ul>");
            currentListLevel++;
        }

        if (isParagraphInsideList) {
            html.append("<li class=\"")
                    .append(paragraph.getStyle())
                    .append("\" style=\"")
                    .append(generateCSSForParagraph(paragraph))
                    .append("\">");
        } else {
            html.append("<p class=\"")
                    .append(paragraph.getStyle())
                    .append("\" style=\"")
                    .append(generateCSSForParagraph(paragraph))
                    .append("\">");
        }

        for (XWPFRun textRegion : paragraph.getRuns()) {
            html.append("<span style='")
                    .append(generateCSSForTextRegion(textRegion))
                    .append("'>")
                    .append(textRegion.text())
                    .append("</span>");

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

        if (isParagraphInsideList) {
            html.append("</li>");
        } else {
            html.append("</p>");
        }

        return html.toString();
    }

    private static String generateCSSForStyles(Set<DocxStyle> styleList) {
        StringBuilder css = new StringBuilder();
        for(DocxStyle style : styleList) {
            css.append(style.toCSS());
        }
        return css.toString();
    }

    private void generateStyleObjects(List<XWPFStyle> usedStyles) {
        for (XWPFStyle style : usedStyles) {
            createAndAddStyleObject(style.getStyleId());
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
        if (ctStyle.getPPr() != null && ctStyle.getPPr().getJc() != null) {
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

    private static String generateCSSForTable(XWPFTable table) {
        StringBuilder css = new StringBuilder();
        css.append("border-collapse: collapse; ");
        switch(table.getWidthType()) {
            // POI docs say: "If the width type is AUTO, DXA, or NIL, the value is 20ths of a point".
            // To translate points into pixels, we multiply by 4/3.
            case DXA:
            case AUTO:
            case NIL:
                css.append("width: ").append(table.getWidth() / 20 * 4 / 3).append("px; ");
                break;
            // "If the width type is PCT, the value is the percentage times 50 (e.g., 2500 for 50%)".
            case PCT:
                css.append("width: ").append(table.getWidth() / 50).append("%; ");
                break;
        }
        return css.toString();
    }

    private static String generateCSSForTableCell(XWPFTableCell cell) {
        StringBuilder css = new StringBuilder();
        CTBorder topBorder = cell.getCTTc().getTcPr().getTcBorders().getTop();
        // We convert line size from half points into points, dividing by two.
        // Then we translate points into pixels, multiplying by 4/3.
        // The result is (size / 2) * 4/3 = size * 4/6 = size * 2/3.
        if (topBorder != null) {
            css.append("border-top-width: ").append(topBorder.getSz().doubleValue() * 2 / 3).append("px; ")
                    .append("border-top-color: #").append(topBorder.xgetColor().getStringValue()).append("; ")
                    .append("border-top-style: ").append(translateBorderValue(topBorder.getVal())).append("; ");
        }
        CTBorder bottomBorder = cell.getCTTc().getTcPr().getTcBorders().getBottom();
        if (bottomBorder != null) {
            css.append("border-bottom-width: ").append(bottomBorder.getSz().doubleValue() * 2 / 3).append("px; ")
                    .append("border-bottom-color: #").append(bottomBorder.xgetColor().getStringValue()).append("; ")
                    .append("border-bottom-style: ").append(translateBorderValue(bottomBorder.getVal())).append("; ");
        }
        CTBorder leftBorder = cell.getCTTc().getTcPr().getTcBorders().getStart();
        if (leftBorder != null) {
            css.append("border-left-width: ").append(leftBorder.getSz().doubleValue() * 2 / 3).append("px; ")
                    .append("border-left-color: #").append(leftBorder.xgetColor().getStringValue()).append("; ")
                    .append("border-left-style: ").append(translateBorderValue(leftBorder.getVal())).append("; ");
        }
        CTBorder rightBorder = cell.getCTTc().getTcPr().getTcBorders().getEnd();
        if (rightBorder != null) {
            css.append("border-right-width: ").append(rightBorder.getSz().doubleValue() * 2 / 3).append("px; ")
                    .append("border-right-color: #").append(rightBorder.xgetColor().getStringValue()).append("; ")
                    .append("border-right-style: ").append(translateBorderValue(rightBorder.getVal())).append(";" );
        }
        System.out.println(css);
        return css.toString();
    }

    private static String translateBorderValue(STBorder.Enum ooxmlValue) {
        switch (ooxmlValue.toString()) {
            case "nil":
            case "none":
                return "none";
            case "single":
                return "solid";
            case "dotted":
                return "dotted";
            case "dashed":
                return "dashed";
            case "double":
                return "double";
            default:
                return "solid";
        }
    }

    @Override
    public String getDefaultExtension() {
        return "html";
    }
}
