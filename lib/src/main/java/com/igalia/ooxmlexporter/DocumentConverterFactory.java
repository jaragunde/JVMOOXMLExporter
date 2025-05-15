package com.igalia.ooxmlexporter;

public class DocumentConverterFactory {
    public static DocumentConverter getConverterForDocument(String inputFile, String outputFile) {
        if (inputFile.endsWith(".pptx")) {
            return new PptxConverter(inputFile, outputFile);
        } else if (inputFile.endsWith(".docx")) {
            return new DocxToHtmlConverter(inputFile, outputFile);
        } else if (inputFile.endsWith(".xls")) {
            return new XlsxConverter(inputFile, outputFile);
        } else if (inputFile.endsWith(".xlsx")) {
            return new XlsxConverter(inputFile, outputFile);
        } else {
            return null;
        }
    }
}
