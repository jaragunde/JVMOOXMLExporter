import de.rototor.pdfbox.graphics2d.PdfBoxGraphics2D;
import de.rototor.pdfbox.graphics2d.PdfBoxGraphics2DFontTextDrawer;
import de.rototor.pdfbox.graphics2d.PdfBoxGraphics2DFontTextDrawerDefaultFonts;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PptxConverter {
    public static void main (String[] args) throws IOException {
        if (args.length == 0) {
            System.out.println("Input file required as argument.");
            System.exit(0);
        }
        String inputFile = args[0];

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(inputFile));
        PDDocument document = new PDDocument();

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
}
