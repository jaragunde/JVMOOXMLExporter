package com.igalia.ooxmlexporter;

import org.apache.poi.examples.ss.html.ToHtml;

import java.io.IOException;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;

public class XlsxConverter implements DocumentConverter {
    private String inputFile;
    private String outputFile;

    public XlsxConverter(String inputFile) {
        this.inputFile = inputFile;
    }

    @Override
    public void convert(String outputFile) {
        this.outputFile = outputFile + "." + getDefaultExtension();
        try (PrintWriter pw = new PrintWriter(this.outputFile, StandardCharsets.UTF_8.name())) {
            ToHtml toHtml = ToHtml.create(inputFile, pw);
            toHtml.setCompleteHTML(true);
            toHtml.printPage();

            ChartConverter chartConverter = new ChartConverter(inputFile);
            chartConverter.convert();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public String getDefaultExtension() {
        return "html";
    }
}
