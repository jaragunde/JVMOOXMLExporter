package com.igalia.ooxmlexporter;

public class DocumentConverterFactory {
    public static DocumentConverter getConverterForDocument(String inputFile, String outputFile) {
        if (inputFile.endsWith(".pptx")) {
            return new PptxConverter(inputFile, outputFile);
        } else if (inputFile.endsWith(".docx")) {
            return new DocxConverter(inputFile, outputFile);
        } else {
            return null;
        }
    }
}
