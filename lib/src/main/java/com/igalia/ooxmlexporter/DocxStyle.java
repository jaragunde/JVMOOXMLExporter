package com.igalia.ooxmlexporter;

import java.math.BigInteger;
import java.util.Objects;

public class DocxStyle {
    private String name;
    private String fontName;
    private BigInteger fontSize;

    public DocxStyle(String name) {
        this.name = name;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public void setFontSize(BigInteger fontSize) {
        this.fontSize = fontSize;
    }

    public String getFontName() {
        return fontName;
    }

    public BigInteger getFontSize() {
        return fontSize;
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
