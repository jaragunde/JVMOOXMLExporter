package com.igalia.ooxmlexporter;

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
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PptxConverter implements DocumentConverter {
    private PDDocument document;
    private String inputFile;
    private String outputFile;

    private static final String SYSTEM_FONT_PATH = "/usr/share/fonts";

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
        if (file.isFile() && file.getName().endsWith(".ttf")) {
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

            for (XSLFSlide slide : ppt.getSlides()) {
                PDPage page = new PDPage(new PDRectangle(ppt.getPageSize().width, ppt.getPageSize().height));
                document.addPage(page);
                PdfBoxGraphics2D pdfBoxGraphics2D = new PdfBoxGraphics2D(document, ppt.getPageSize().width, ppt.getPageSize().height);

                pdfBoxGraphics2D.setFontTextDrawer(createFontTextDrawer());
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
