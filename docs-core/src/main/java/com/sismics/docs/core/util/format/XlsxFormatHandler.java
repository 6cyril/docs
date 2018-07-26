package com.sismics.docs.core.util.format;

import com.google.common.io.Closer;
import com.sismics.util.mime.MimeType;
import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.pdmodel.PDDocument;
//import org.apache.pdfbox.pdmodel.PDPage;
//import org.apache.pdfbox.pdmodel.PDPageContentStream;
//import org.apache.pdfbox.pdmodel.common.PDRectangle;
//import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
//import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
//import org.apache.poi.xslf.extractor.XSLFPowerPointExtractor;
//import org.apache.poi.xslf.usermodel.XMLSlideShow;
//import org.apache.poi.xslf.usermodel.XSLFSlide;

import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;

//import java.awt.*;
//import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * Xlsx format handler.
 *
 * @author bgamard
 */
public class XlsxFormatHandler implements FormatHandler {
    /**
     * Cached XLSX loaded file.
     */
    //private XMLSlideShow slideShow;
    private XSSFWorkbook workBook;

    @Override
    public boolean accept(String mimeType) {
        return MimeType.OFFICE_SHEET.equals(mimeType);
    }

    @Override
    public BufferedImage generateThumbnail(Path file) throws Exception {
/*        XMLSlideShow pptx = loadPPtxFile(file);
        if (pptx.getSlides().size() > 0) {
            return generateImageFromSlide(pptx, 0);
        }
*/
        return null;
    }

//    @Override
//    public String extractContent(String language, Path file) throws Exception {
//        XMLSlideShow pptx = loadPPtxFile(file);
//        return new XSLFPowerPointExtractor(pptx).getText();
//    }

    @Override
    public String extractContent(String language, Path file) throws Exception {
//          OPCPackage pkg = OPCPackage.open( file.toString() );
//          XSSFWorkbook wb = new XSSFWorkbook(pkg);
          XSSFWorkbook wb = loadXLsxFile(file);
          XSSFExcelExtractor ext = new XSSFExcelExtractor(wb);
          return ext.getText();
    }


    @Override
    public void appendToPdf(Path file, PDDocument doc, boolean fitImageToPage, int margin, MemoryUsageSetting memUsageSettings, Closer closer) throws Exception {
/*        XMLSlideShow pptx = loadPPtxFile(file);
        List<XSLFSlide> slides = pptx.getSlides();
        Dimension pgsize = pptx.getPageSize();
        for (int slideIndex = 0; slideIndex < slides.size(); slideIndex++) {
            // One PDF page per slide
            PDPage page = new PDPage(new PDRectangle(pgsize.width, pgsize.height));
            try (PDPageContentStream contentStream = new PDPageContentStream(doc, page)) {
                BufferedImage bim = generateImageFromSlide(pptx, slideIndex);
                PDImageXObject pdImage = LosslessFactory.createFromImage(doc, bim);
                contentStream.drawImage(pdImage, 0, page.getMediaBox().getHeight() - pdImage.getHeight());
            }
            doc.addPage(page);
        }*/
    }

    public XSSFWorkbook loadXLsxFile(Path file) throws Exception {
        if  (workBook == null) {
            try (OPCPackage pkg = OPCPackage.open( file.toString())){
                workBook = new XSSFWorkbook(pkg);
            }
        }
        return workBook;
    }

/*
    private XMLSlideShow loadPPtxFile(Path file) throws Exception {
        if (slideShow == null) {
            try (InputStream inputStream = Files.newInputStream(file)) {
                slideShow = new XMLSlideShow(inputStream);
            }
        }
        return slideShow;
    }
*/


    /**
     * Generate an image from a PPTX slide.
     *
     * @param pptx PPTX
     * @param slideIndex Slide index
     * @return Image
     */
/*    private BufferedImage generateImageFromSlide(XMLSlideShow pptx, int slideIndex) {
        Dimension pgsize = pptx.getPageSize();
        BufferedImage img = new BufferedImage(pgsize.width, pgsize.height,BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics = img.createGraphics();
        graphics.setPaint(Color.white);
        graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
        pptx.getSlides().get(slideIndex).draw(graphics);
        return img;
    }
*/

}

