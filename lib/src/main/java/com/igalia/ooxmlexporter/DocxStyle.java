package com.igalia.ooxmlexporter;

import java.util.Objects;

public class DocxStyle {
    private String name;
    private FontFormat fontFormat;

    public DocxStyle(String name, DocxStyle copyFrom) {
        this.name = name;
        this.fontFormat = new FontFormat();
        this.fontFormat.setFontName(copyFrom.fontFormat.getFontName());
        this.fontFormat.setFontColor(copyFrom.fontFormat.getFontColor());
        this.fontFormat.setFontSize(copyFrom.fontFormat.getFontSize());
        this.fontFormat.setBold(copyFrom.fontFormat.getBold());
        this.fontFormat.setItalic(copyFrom.fontFormat.getItalic());
    }

    public DocxStyle(String name) {
        this.name = name;
        this.fontFormat = new FontFormat();
    }

    public String getName() {
        return name;
    }

    public FontFormat getFontFormat() {
        return fontFormat;
    }

    public String toCSS() {
        StringBuilder css = new StringBuilder();
        css.append(".").append(name).append(" {")
                .append(fontFormat.toCSS())
                .append("} ");
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
