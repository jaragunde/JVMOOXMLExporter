package com.igalia.ooxmlexporter;

public class DocumentConverterFactory {

    public static enum DocumentFormat {
        PDF,
        HTML
    }

    public static DocumentConverter getConverterForDocument(String inputFile, DocumentFormat format) {
        if (inputFile.endsWith(".pptx") && format == DocumentFormat.PDF) {
            return new PptxConverter(inputFile);
        } else if (inputFile.endsWith(".docx") && format == DocumentFormat.PDF) {
            return new DocxConverter(inputFile);
        } else if (inputFile.endsWith(".docx") && format == DocumentFormat.HTML) {
            return new DocxToHtmlConverter(inputFile);
        } else if (inputFile.endsWith(".xls") && format == DocumentFormat.HTML) {
            return new XlsxConverter(inputFile);
        } else if (inputFile.endsWith(".xlsx") && format == DocumentFormat.HTML) {
            return new XlsxConverter(inputFile);
        } else {
            return null;
        }
    }
}
