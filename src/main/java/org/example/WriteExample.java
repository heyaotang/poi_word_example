package org.example;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigInteger;

import javax.imageio.ImageIO;

import org.apache.commons.io.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WriteExample {

    private static final int MAX_IMAGE_WIDTH = 388;
    private static final String TEMPLATE_PATH = "F:\\Example.docx";
    private static final String IMAGE_PATH = "F:\\JPG.jpg";
    private static final String EXPORT_PATH = "F:\\";

    private void writeHeader(XWPFDocument doc, String text, Integer paragraphIndex) {
        String styleId = doc.getStyles().getStyleWithName("heading 1").getStyleId();
        XWPFParagraph para = null;
        if (paragraphIndex != null) {
            para = doc.getParagraphs().get(paragraphIndex);
        } else {
            para = doc.createParagraph();
        }
        para.setStyle(styleId);
        XWPFRun run = para.createRun();
        run.setText(text);
    }

    private void writeText(XWPFDocument doc, String text, String numID) {
        String styleId = doc.getStyles().getStyleWithName("Normal").getStyleId();
        XWPFParagraph para = doc.createParagraph();
        para.setStyle(styleId);
        if (numID != null) {
            para.setNumID(new BigInteger(numID));
        }
        XWPFRun run = para.createRun();
        run.setText(text);
    }


    public void writeToDocx() throws Exception {
        XWPFDocument docx = new XWPFDocument(new FileInputStream(TEMPLATE_PATH));
        //write first header
        writeHeader(docx, "KEYWORD_1", 0);
        writeText(docx, "", null);

        //write other header
        writeHeader(docx, "KEYWORD_2", null);
        writeText(docx, "", null);

        //write normal text
        writeText(docx, "The Apache POI team is pleased to announce the release of 5.2.2. Several dependencies were updated to their latest versions to pick up security fixes and other improvements.", null);
        writeText(docx, "", null);

        //write normal text with order code
        writeText(docx, "The Apache POI team.", "1");
        writeText(docx, "The Apache POI team.", "1");
        writeText(docx, "The Apache POI team.", "1");
        writeText(docx, "The Apache POI team.", "1");
        writeText(docx, "", null);

        //write normal text with order code
        writeText(docx, "The Apache POI team.", "2");
        writeText(docx, "The Apache POI team.", "2");
        writeText(docx, "The Apache POI team.", "2");
        writeText(docx, "The Apache POI team.", "2");
        writeText(docx, "", null);

        //write image
        File imageFile = new File(IMAGE_PATH);
        String name = imageFile.getName();

        BufferedImage sourceImg = ImageIO.read(imageFile);
        int width = sourceImg.getWidth();
        int height = sourceImg.getHeight();
        if (width > MAX_IMAGE_WIDTH) {
            height = height * MAX_IMAGE_WIDTH / width;
            width = MAX_IMAGE_WIDTH;
        }

        FileInputStream input = new FileInputStream(imageFile);
        docx.createParagraph().createRun().addPicture(input, Document.PICTURE_TYPE_JPEG, name, Units.toEMU(width), Units.toEMU(height));

        OutputStream os = new FileOutputStream(EXPORT_PATH + "Export_" + System.currentTimeMillis() + ".docx");
        docx.write(os);

        IOUtils.close(input);
        IOUtils.close(docx);
        IOUtils.close(os);
    }
}
