package com.igalia.ooxmlexporter;

import de.rototor.pdfbox.graphics2d.PdfBoxGraphics2D;
import de.rototor.pdfbox.graphics2d.PdfBoxGraphics2DFontTextDrawer;
import de.rototor.pdfbox.graphics2d.PdfBoxGraphics2DFontTextDrawerDefaultFonts;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PptxConverter implements DocumentConverter {
    private PDDocument document;
    private String inputFile;
    private String outputFile;

    private static final String SYSTEM_FONT_PATH = "/usr/share/fonts";
    private static final List<String> FONT_DENYLIST = Arrays.asList("NotoColorEmoji.ttf");

    public PptxConverter(String inputFile, String outputFile) {
        this.inputFile = inputFile;
        this.outputFile = outputFile;
    }

    private PdfBoxGraphics2DFontTextDrawer createFontTextDrawer() {
        PdfBoxGraphics2DFontTextDrawerDefaultFonts fontTextDrawer = new PdfBoxGraphics2DFontTextDrawerDefaultFonts();
        addFontsToDrawerRecursively(new File(SYSTEM_FONT_PATH), fontTextDrawer);
        return fontTextDrawer;
    }

    private void addFontsToDrawerRecursively(File file, PdfBoxGraphics2DFontTextDrawer fontTextDrawer) {
        if (file.isFile() && file.getName().endsWith(".ttf") &&
                !FONT_DENYLIST.contains(file.getName())) {
            fontTextDrawer.registerFont(file);
        }
        if (file.isDirectory()) {
            for (File innerFile : file.listFiles()) {
                addFontsToDrawerRecursively(innerFile, fontTextDrawer);
            }
        }
    }

    public void convert() {
        try {
            XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(inputFile));
            document = new PDDocument();
            PdfBoxGraphics2DFontTextDrawer fontTextDrawer = createFontTextDrawer();

            for (XSLFSlide slide : ppt.getSlides()) {
                try {
                    convertSlide(ppt, fontTextDrawer, slide);
                } catch (RuntimeException e) {
                    // Exceptions can be caused by unsupported fonts, but we only learn this
                    // in the rendering phase, not in the font loading phase. In case of
                    // problems, we try once again with a reset font text drawer
                    convertSlide(ppt, new PdfBoxGraphics2DFontTextDrawerDefaultFonts(), slide);
                }
            }
            document.save(new File(outputFile));
            document.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void convertSlide(XMLSlideShow ppt, PdfBoxGraphics2DFontTextDrawer fontTextDrawer, XSLFSlide slide) throws IOException {
        PDPage page = new PDPage(new PDRectangle(ppt.getPageSize().width, ppt.getPageSize().height));
        document.addPage(page);
        PdfBoxGraphics2D pdfBoxGraphics2D = new PdfBoxGraphics2D(document, ppt.getPageSize().width, ppt.getPageSize().height);

        pdfBoxGraphics2D.setFontTextDrawer(fontTextDrawer);
        slide.draw(pdfBoxGraphics2D);
        pdfBoxGraphics2D.dispose();
        PDPageContentStream contentStream = new PDPageContentStream(document, page);
        contentStream.drawForm(pdfBoxGraphics2D.getXFormObject());
        contentStream.close();
    }
}
