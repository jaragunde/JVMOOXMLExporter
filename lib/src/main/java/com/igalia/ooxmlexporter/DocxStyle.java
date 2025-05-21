package com.igalia.ooxmlexporter;

import java.math.BigInteger;
import java.util.Objects;

public class DocxStyle {
    private String name;
    private String fontName;
    private String fontColor;
    private BigInteger fontSize;
    private Boolean bold;
    private Boolean italic;

    public DocxStyle(String name, DocxStyle copyFrom) {
        this.name = name;
        this.fontName = copyFrom.fontName;
        this.fontSize = copyFrom.fontSize;
        this.fontColor = copyFrom.fontColor;
        this.bold = copyFrom.bold;
        this.italic = copyFrom.italic;
    }

    public DocxStyle(String name) {
        this.name = name;
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

    public String getName() {
        return name;
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
        css.append(".").append(name).append(" {");
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
        css.append("} ");
        return css.toString();
    }

    @Override
    public boolean equals(Object o) {
        if (o == null || getClass() != o.getClass()) return false;
        DocxStyle docxStyle = (DocxStyle) o;
        return Objects.equals(name, docxStyle.name);
    }

    @Override
    public int hashCode() {
        return Objects.hashCode(name);
    }
}
