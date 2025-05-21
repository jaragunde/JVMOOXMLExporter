package com.igalia.ooxmlexporter;

import java.math.BigInteger;

public class FontFormat {
    private String fontName;
    private String fontColor;
    private BigInteger fontSize;
    private Boolean bold;
    private Boolean italic;

    public FontFormat() { }

    public FontFormat(String fontName, String fontColor, BigInteger fontSize, Boolean bold, Boolean italic) {
        this.fontName = fontName;
        this.fontColor = fontColor;
        this.fontSize = fontSize;
        this.bold = bold;
        this.italic = italic;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public void setFontSize(BigInteger fontSize) {
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

    public String getFontName() {
        return fontName;
    }

    public BigInteger getFontSize() {
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

    public String toCSS() {
        StringBuilder css = new StringBuilder();
        if (fontName != null) {
            css.append("font-family: \"").append(fontName).append("\"; ");
        }
        if (fontSize != null) {
            css.append("font-size: ").append(fontSize).append("px; ");
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
        return css.toString();
    }
}
