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
    public static void main (String args[]) throws IOException {
        //Class c = Class.forName("PptxConverter");
        XMLSlideShow ppt = new XMLSlideShow(PptxConverter.class.getResourceAsStream("fonts-in-text.pptx"));
        PDDocument document = new PDDocument();

        for (XSLFSlide slide : ppt.getSlides()) {
            PDPage page = new PDPage(new PDRectangle(ppt.getPageSize().width, ppt.getPageSize().height));
            document.addPage(page);
            PdfBoxGraphics2D pdfBoxGraphics2D = new PdfBoxGraphics2D(document, ppt.getPageSize().width, ppt.getPageSize().height);

            pdfBoxGraphics2D.setFontTextDrawer(new PdfBoxGraphics2DFontTextDrawerDefaultFonts());
            slide.draw(pdfBoxGraphics2D);
            pdfBoxGraphics2D.dispose();

            PDPageContentStream contentStream = new PDPageContentStream(document, page);
            contentStream.drawForm(pdfBoxGraphics2D.getXFormObject());
            contentStream.close();
        }
        document.save(new File("output-PptToPdf.pdf"));
        document.close();
    }
}
