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
    private PDDocument document;
    private String inputFile;
    private String outputFile;

    public PptxConverter(String inputFile, String outputFile) {
        this.inputFile = inputFile;
        this.outputFile = outputFile;
    }

    public void convert() {
        try {
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
            document.save(new File(outputFile));
            document.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
