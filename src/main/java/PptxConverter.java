import de.rototor.pdfbox.graphics2d.PdfBoxGraphics2D;
import de.rototor.pdfbox.graphics2d.PdfBoxGraphics2DFontTextDrawerDefaultFonts;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xwpf.usermodel.*;

public class PptxConverter {
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

    public static void convertDocx(String inputFile) throws IOException {
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
        document.save(new File(inputFile + ".pdf"));
        document.close();
    }

    public static void convertPptx(String inputFile) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(inputFile));
        document = new PDDocument();

        for (XSLFSlide slide : ppt.getSlides()) {
            PDPage page = new PDPage(new PDRectangle(ppt.getPageSize().width, ppt.getPageSize().height));
            document.addPage(page);
            PdfBoxGraphics2D pdfBoxGraphics2D = new PdfBoxGraphics2D(document, ppt.getPageSize().width, ppt.getPageSize().height);

            PdfBoxGraphics2DFontTextDrawerDefaultFonts fontTextDrawer = new PdfBoxGraphics2DFontTextDrawerDefaultFonts();
            fontTextDrawer.registerFont(new File("/usr/share/fonts/liberation-serif/LiberationSerif-Regular.ttf"));
            fontTextDrawer.registerFont(new File("/usr/share/fonts/google-carlito-fonts/Carlito-Regular.ttf"));
            pdfBoxGraphics2D.setFontTextDrawer(fontTextDrawer);
            slide.draw(pdfBoxGraphics2D);
            pdfBoxGraphics2D.dispose();

            PDPageContentStream contentStream = new PDPageContentStream(document, page);
            contentStream.drawForm(pdfBoxGraphics2D.getXFormObject());
            contentStream.close();
        }
        document.save(new File(inputFile + ".pdf"));
        document.close();
    }

    public static void main (String[] args) throws IOException {
        if (args.length == 0) {
            System.out.println("Input file required as argument.");
            System.exit(0);
        }
        String inputFile = args[0];

        if (inputFile.endsWith(".pptx")) {
            convertPptx(inputFile);
        } else if (inputFile.endsWith(".docx")) {
            convertDocx(inputFile);
        } else {
            System.out.println("File type unsupported.");
            System.exit(0);
        }
    }
}
