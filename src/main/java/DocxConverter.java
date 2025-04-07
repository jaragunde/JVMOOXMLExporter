import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class DocxConverter {
    public static PDDocument document;

    public static PDFont defaultFont() {
        return new PDType1Font(Standard14Fonts.FontName.TIMES_ROMAN);
    }

    public static PDFont fontMatch(String fontName) throws IOException {
        if (fontName == null) {
            return defaultFont();
        }
        switch (fontName) {
            case "Cantallel":
                return defaultFont();
            case "Nimbus Mono PS":
                return defaultFont();
            case "Carlito":
                return PDType0Font.load(document, new File("/usr/share/fonts/google-carlito-fonts/Carlito-Regular.ttf"));
            case "Liberation Serif":
                return PDType0Font.load(document, new File("/usr/share/fonts/liberation-serif/LiberationSerif-Regular.ttf"));
            default:
                return defaultFont();
        }
    }

    public static void convertDocx(String inputFile, String outputFile) throws IOException {
        XWPFDocument docx = new XWPFDocument(new FileInputStream(inputFile));
        document = new PDDocument();
        PDPage page = new PDPage();
        document.addPage(page);
        PDPageContentStream contentStream = new PDPageContentStream(document, page);
        int textLocation = 0;

        for (IBodyElement bodyElement : docx.getBodyElements()) {
            if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                for (XWPFRun textRegion : ((XWPFParagraph) bodyElement).getRuns()) {
                    String text = textRegion.text();
                    String fontName = textRegion.getFontName();
                    System.out.println(text);
                    System.out.println(fontName);
                    textLocation += 25;

                    contentStream.beginText();
                    contentStream.setFont(fontMatch(fontName), 12);
                    contentStream.newLineAtOffset(25, textLocation);
                    contentStream.showText(text);
                    contentStream.endText();
                }
            }
        }
        contentStream.close();
        document.save(new File(outputFile));
        document.close();
    }
}
