package com.igalia.ooxmlexporter;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

import java.math.BigInteger;

public class FontFormat {
    private String fontName;
    private String fontColor;
    private Double fontSize;  // Font size is expressed in points.
    private Boolean bold;
    private Boolean italic;
    private ParagraphAlignment alignment;

    public FontFormat() { }

    public FontFormat(String fontName, String fontColor, Double fontSize, Boolean bold, Boolean italic) {
        this.fontName = fontName;
        this.fontColor = fontColor;
        this.fontSize = fontSize;
        this.bold = bold;
        this.italic = italic;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public void setFontSize(Double fontSize) {
        this.fontSize = fontSize;
    }

    public void setFontColor(String fontColor) {
        this.fontColor = fontColor;
    }

    public void setBold(Boolean bold) {
        this.bold = bold;
    }

    public void setItalic(Boolean italic) {
        this.italic = italic;
    }

    public void setParagraphAlignment(ParagraphAlignment alignment) {
        this.alignment = alignment;
    }

    public String getFontName() {
        return fontName;
    }

    public Double getFontSize() {
        return fontSize;
    }

    public String getFontColor() {
        return fontColor;
    }

    public Boolean getBold() {
        return bold;
    }

    public Boolean getItalic() {
        return italic;
    }

    public ParagraphAlignment getAlignment() {
        return alignment;
    }

    public String toCSS() {
        StringBuilder css = new StringBuilder();
        if (fontName != null) {
            css.append("font-family: \"").append(fontName).append("\"; ");
        }
        if (fontSize != null) {
            // We multiply font size by 4/3 to translate points into pixels.
            double sizeInPx = fontSize * 4 / 3;
            css.append("font-size: ").append(sizeInPx).append("px; ");
        }
        if (bold != null && bold) {
            css.append("font-weight: bold; ");
        }
        if (italic != null && italic) {
            css.append("font-style: italic; ");
        }
        if (fontColor != null) {
            css.append("color: #").append(fontColor).append("; ");
        }
        if (alignment != null) {
            switch (alignment) {
                case ParagraphAlignment.LEFT:
                case ParagraphAlignment.START:
                    css.append("text-align: left;");
                    break;
                case ParagraphAlignment.RIGHT:
                case ParagraphAlignment.END:
                    css.append("text-align: right;");
                    break;
                case ParagraphAlignment.CENTER:
                    css.append("text-align: center;");
                    break;
                case ParagraphAlignment.BOTH:
                    css.append("text-align: justify;");
                    break;
            }
        }
        return css.toString();
    }
}
