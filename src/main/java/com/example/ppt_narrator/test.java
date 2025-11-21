package com.example.ppt_narrator;

import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

public class test {
    public static void main(String[] args) throws Exception {

        String emojis = getSpeakerNotes();
        // System.out.println(emojis);
        System.out.println(removeEmojis(emojis));
        // for (int ctr = 0; ctr < emojis.length(); ctr++) {
        // System.out.println((int) emojis.charAt(ctr));
        // }
    }

    private static String removeEmojis(String text) {
        StringBuilder sb = new StringBuilder(text.length());
        char c;
        for (int ctr = 0; ctr < text.length();) {
            c = text.charAt(ctr);
            if (c > 127) {
                c = text.charAt(++ctr);
                while (!(Character.isAlphabetic(c) || Character.isDigit(c) || Character.isWhitespace(c)))
                    c = text.charAt(++ctr);
                if (Character.isDigit(c))
                    ctr++;
            } else {
                sb.append(c);
                ctr++;
            }
        }
        return sb.toString();
    }

    private static boolean isEmoji(char c) {
        return ((int) c) > 127;

    }

    static String getSpeakerNotes() throws Exception {
        try (FileInputStream fis = new FileInputStream("D:\\Learning\\ppt_to_audios\\copy.pptx");
                XMLSlideShow ppt = new XMLSlideShow(OPCPackage.open(fis))) {

            int slideNum = 1;

            for (XSLFSlide slide : ppt.getSlides()) {
                return extractNotes(slide);
            }

        }
        return null;
    }

    private static String extractNotes(XSLFSlide slide) {
        StringBuilder sb = new StringBuilder(1024);
        int emojiCount = 0;
        XSLFNotes notes = slide.getNotes();
        int totTxtRuns = 0;
        if (notes != null) {
            for (XSLFShape shape : notes.getShapes()) {
                if (shape instanceof XSLFTextShape ts) {
                    sb.append(ts.getText()).append(' ');
                    // List<XSLFTextParagraph> paras = ts.getTextParagraphs();
                    // for (XSLFTextParagraph para : paras) {
                    // List<XSLFTextRun> txtRuns = para.getTextRuns();
                    // totTxtRuns += txtRuns.size();
                    // for (XSLFTextRun txtRun : txtRuns) {
                    // if (!isEmoji(txtRun)) {
                    // sb.append(txtRun.getRawText());
                    // }
                    // }
                    // sb.append("\n");
                    // }
                }
            }
            // System.out.println(totTxtRuns);
        }
        return sb.toString().trim();
    }

    static boolean isEmoji(XSLFTextRun txtRun) {
        String rawText = txtRun.getRawText();
        for (char c : rawText.toCharArray()) {
            if (c > 1000) {
                return true;
            }
        }
        return false;
    }

    static void printRaw(XSLFTextRun txtRun) {
        String rawText = txtRun.getRawText();
        for (char c : rawText.toCharArray()) {
            System.out.print("" + ((int) c) + ":");
        }
        System.out.println();
    }
}
