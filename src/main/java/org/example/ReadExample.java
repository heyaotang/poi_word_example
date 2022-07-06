package org.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.io.IOUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ReadExample {

    private static final String DOC_PATH = "F:\\Export.doc";
    private static final String DOCX_PATH = "F:\\Export.docx";
    private static final String EXPORT_PATH = "F:\\";

    public void parseDoc() {
        try (FileInputStream fis = new FileInputStream(DOC_PATH); HWPFDocument document = new HWPFDocument(fis);) {
            PicturesSource pictures = new PicturesSource(document);
            PicturesTable pictureTable = document.getPicturesTable();
            Range r = document.getRange();
            for (int i = 0; i < r.numParagraphs(); i++) {
                Paragraph p = r.getParagraph(i);

                // for text
                String paragraphText = p.text();
                //具体的业务逻辑
                System.out.println("paragraphText => " + paragraphText);

                // for picture
                for (int j = 0; j < p.numCharacterRuns(); j++) {
                    CharacterRun cr = p.getCharacterRun(j);
                    if (pictureTable.hasPicture(cr)) {
                        Picture picture = pictures.getFor(cr);
                        String path = EXPORT_PATH + System.currentTimeMillis() + ".jpg";
                        FileOutputStream fos = new FileOutputStream(path);
                        picture.writeImageContent(fos);
                        IOUtils.close(fos);
                        System.out.println("Picture: " + picture.getMimeType() + "\t" + path);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void parseDocx() {
        try (FileInputStream fis = new FileInputStream(DOCX_PATH); XWPFDocument document = new XWPFDocument(fis);) {
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (int i = 0; i < paragraphs.size(); i++) {
                XWPFParagraph p = paragraphs.get(i);

                // for text
                String paragraphText = p.getParagraphText();
                System.out.println("paragraphs => " + paragraphText);

                // for picture
                List<XWPFRun> runs = p.getRuns();
                for (XWPFRun run : runs) {
                    List<XWPFPicture> pictures = run.getEmbeddedPictures();
                    for (XWPFPicture picture : pictures) {
                        XWPFPictureData pictureData = picture.getPictureData();
                        String path = EXPORT_PATH + System.currentTimeMillis() + ".jpg";
                        FileOutputStream fos = new FileOutputStream(path);
                        IOUtils.write(pictureData.getData(), fos);
                        IOUtils.close(fos);
                        System.out.println("Picture: " + path);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}


//这个类是poi官网找的
class PicturesSource {

    private PicturesTable picturesTable;
    private Set<Picture> output = new HashSet<Picture>();
    private Map<Integer, Picture> lookup;
    private List<Picture> nonU1based;
    private List<Picture> all;
    private int pn = 0;

    public PicturesSource(HWPFDocument doc) {
        picturesTable = doc.getPicturesTable();
        all = picturesTable.getAllPictures();

        // Build the Offset-Picture lookup map
        lookup = new HashMap<Integer, Picture>();
        for (Picture p : all) {
            lookup.put(p.getStartOffset(), p);
        }

        // Work out which Pictures aren't referenced by
        //  a \u0001 in the main text
        // These are \u0008 escher floating ones, ones
        //  found outside the normal text, and who
        //  knows what else...
        nonU1based = new ArrayList<Picture>();
        nonU1based.addAll(all);
        Range r = doc.getRange();
        for (int i = 0; i < r.numCharacterRuns(); i++) {
            CharacterRun cr = r.getCharacterRun(i);
            if (picturesTable.hasPicture(cr)) {
                Picture p = getFor(cr);
                int at = nonU1based.indexOf(p);
                nonU1based.set(at, null);
            }
        }
    }

    private boolean hasPicture(CharacterRun cr) {
        return picturesTable.hasPicture(cr);
    }

    private void recordOutput(Picture picture) {
        output.add(picture);
    }

    private boolean hasOutput(Picture picture) {
        return output.contains(picture);
    }

    private int pictureNumber(Picture picture) {
        return all.indexOf(picture) + 1;
    }

    public Picture getFor(CharacterRun cr) {
        return lookup.get(cr.getPicOffset());
    }

    /**
     * Return the next unclaimed one, used towards the end
     */
    private Picture nextUnclaimed() {
        Picture p = null;
        while (pn < nonU1based.size()) {
            p = nonU1based.get(pn);
            pn++;
            if (p != null) {
                return p;
            }
        }
        return null;
    }
}