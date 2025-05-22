package com.igalia.ooxmlexporter;

public class DocumentConverterFactory {

    public static enum DocumentFormat {
        PDF,
        HTML
    }

    public static DocumentConverter getConverterForDocument(String inputFile, DocumentFormat format) {
        if (inputFile.toLowerCase().endsWith(".pptx") && format == DocumentFormat.PDF) {
            return new PptxConverter(inputFile);
        } else if (inputFile.toLowerCase().endsWith(".docx") && format == DocumentFormat.PDF) {
            return new DocxConverter(inputFile);
        } else if (inputFile.toLowerCase().endsWith(".docx") && format == DocumentFormat.HTML) {
            return new DocxToHtmlConverter(inputFile);
        } else if (inputFile.toLowerCase().endsWith(".xls") && format == DocumentFormat.HTML) {
            return new XlsxConverter(inputFile);
        } else if (inputFile.toLowerCase().endsWith(".xlsx") && format == DocumentFormat.HTML) {
            return new XlsxConverter(inputFile);
        } else {
            return null;
        }
    }
}
