package com.igalia.ooxmlexporter;

import org.apache.poi.examples.ss.html.ToHtml;

import java.io.IOException;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;

public class XlsxConverter implements DocumentConverter {
    private String inputFile;
    private String outputFile;

    public XlsxConverter(String inputFile, String outputFile) {
        this.inputFile = inputFile;
        this.outputFile = outputFile;
    }

    @Override
    public void convert() {
        try (PrintWriter pw = new PrintWriter(outputFile, StandardCharsets.UTF_8.name())) {
            ToHtml toHtml = ToHtml.create(inputFile, pw);
            toHtml.setCompleteHTML(true);
            toHtml.printPage();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
