package com.example.ppt_narrator;

import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.xslf.usermodel.*;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.beans.factory.annotation.Autowired;

import software.amazon.awssdk.core.ResponseInputStream;
import software.amazon.awssdk.services.polly.PollyClient;
import software.amazon.awssdk.services.polly.model.*;

// import com.mpatric.mp3agic.Mp3File;

import java.io.*;
// import java.nio.charset.StandardCharsets;
import java.nio.file.*;

@SpringBootApplication
public class PptNarratorApplication implements CommandLineRunner {

    @Autowired
    private PollyClient polly;

    // 👉 To keep count of total characters across all slides
    private int totalCharacters = 0;

    public static void main(String[] args) {
        SpringApplication.run(PptNarratorApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {

        if (args.length == 0) {
            System.err.println(
                    "❌ Usage: mvn spring-boot:run -Dspring-boot.run.arguments=\"<path-to-pptx>[output-directory]\"");
            System.exit(1);
        }

        String pptxFile = args[0];
        String outputDir = args.length > 1 ? args[1] : "slide_audios";

        // Extract PPT filename without .pptx
        String pptName = Paths.get(pptxFile).getFileName().toString().replace(".pptx", "");
        System.out.println("📂 PPT Name: " + pptName);

        System.out.println("📂 Input PPTX: " + pptxFile);
        System.out.println("🎧 Output Directory: " + outputDir);

        Files.createDirectories(Paths.get(outputDir));

        // String outputPptx = Paths
        // .get(outputDir, Paths.get(pptxFile).getFileName().toString().replace(".pptx",
        // "_with_audio.pptx"))
        // .toString();
        // Files.copy(Paths.get(pptxFile), Paths.get(outputPptx),
        // StandardCopyOption.REPLACE_EXISTING);

        // try (FileInputStream fis = new FileInputStream(pptxFile);
        // XMLSlideShow ppt = new XMLSlideShow(OPCPackage.open(fis))) {

        // int slideNum = 1;

        // for (XSLFSlide slide : ppt.getSlides()) {
        // String notesText = extractNotes(slide);
        // notesText = cleanNotes(notesText);
        // // 👉 Count characters
        // countCharacters(slideNum, notesText);
        // slideNum++;
        // }

        // }
        // System.out.println("Total Characters : " + totalCharacters);
        // }

        try (FileInputStream fis = new FileInputStream(pptxFile);
                XMLSlideShow ppt = new XMLSlideShow(OPCPackage.open(fis))) {

            int slideNum = 1;

            for (XSLFSlide slide : ppt.getSlides()) {
                // if (slideNum <= 5) {
                // slideNum++;
                // continue;
                // }
                String notesText = extractNotes(slide);

                // 👉 Count characters
                int count = countCharacters(slideNum, notesText);
                if (count >= 2999) {
                    System.out.println("Skipping Slide " + slideNum + " has " + count + " characters");
                    slideNum++;
                    continue;
                }

                String cleanedNotes = cleanNotes(notesText);

                if (cleanedNotes.isBlank()) {
                    System.out.println("Slide " + slideNum + ": No notes found, skipping audio.");
                    slideNum++;
                    continue;
                }

                String audioFilePath = outputDir + "/" + pptName + "_" + slideNum + ".mp3";
                Path audioPath = Paths.get(audioFilePath);

                if (!Files.exists(audioPath)) {

                    SynthesizeSpeechRequest request = SynthesizeSpeechRequest.builder()
                            .text(cleanedNotes)
                            .voiceId("Danielle")
                            .textType(TextType.SSML)
                            .engine("generative")
                            .outputFormat(OutputFormat.MP3)
                            .build();

                    try (ResponseInputStream<SynthesizeSpeechResponse> stream = polly.synthesizeSpeech(request);
                            OutputStream out = new FileOutputStream(audioFilePath)) {

                        stream.transferTo(out);
                    }

                    System.out.println("✅ Generated audio: " + audioFilePath);

                } else {
                    System.out.println("ℹ️ Using existing audio: " + audioFilePath);
                }

                // ---------------------------------------------
                // 📌 Embed audio in ppt/media folder
                // ---------------------------------------------
                // OPCPackage pkg = ppt.getPackage();

                // String mediaPath = "/ppt/media/audio" + pptName + "_" + slideNum + ".mp3";
                // PackagePartName partName = PackagingURIHelper.createPartName(mediaPath);

                // if (pkg.containPart(partName)) {
                // pkg.removePart(partName);
                // }

                // PackagePart audioPart = pkg.createPart(partName, "audio/mpeg");

                // try (InputStream is = Files.newInputStream(audioPath);
                // OutputStream os = audioPart.getOutputStream()) {

                // is.transferTo(os);
                // }

                // String relType =
                // "http://schemas.openxmlformats.org/officeDocument/2006/relationships/media";

                // PackageRelationship rel = slide.getPackagePart()
                // .addRelationship(partName, TargetMode.INTERNAL, relType);

                // String relId = rel.getId();

                // System.out.println("🔊 Embedded audio with relId: " + relId);

                // // ---------------------------------------------
                // // ⏱ Extract MP3 duration (to set auto-advance)
                // // ---------------------------------------------
                // Mp3File mp3 = new Mp3File(audioFilePath);
                // int durationSeconds = (int) Math.ceil(mp3.getLengthInSeconds());
                // int advTimeMs = durationSeconds * 1000;

                // System.out.println("⏱ Duration: " + durationSeconds + " seconds");

                // // ---------------------------------------------
                // // 📝 Modify slide XML to insert audio + timings
                // // ---------------------------------------------
                // PackagePart slidePart = slide.getPackagePart();
                // String slideXml;

                // try (InputStream slideStream = slidePart.getInputStream()) {
                // slideXml = new String(slideStream.readAllBytes(), StandardCharsets.UTF_8);
                // }

                // String audioXml =
                // "<p:audioCdxmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\""
                // +
                // "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                // + "<p:st r:embed=\"" + relId + "\"/>"
                // + "</p:audioCdxmlns:p=>\n";

                // // Insert before </p:spTree>
                // if (slideXml.contains("</p:spTree>")) {
                // slideXml = slideXml.replace("</p:spTree>", audioXml + "</p:spTree>");
                // } else {
                // slideXml += "\n" + audioXml;
                // }

                // // Insert transition timing
                // if (slideXml.contains("<p:transition")) {

                // slideXml = slideXml.replaceFirst(
                // "<p:transition([^>]*)>",
                // "<p:transition$1 advTm=\"" + advTimeMs + "\">");

                // } else {

                // String marker = "</p:cSld>";

                // if (slideXml.contains(marker)) {
                // slideXml = slideXml.replace(
                // marker,
                // "<p:transition advTm=\"" + advTimeMs + "\"/>" + marker);
                // }
                // }

                // try (OutputStream out = slidePart.getOutputStream()) {
                // out.write(slideXml.getBytes(StandardCharsets.UTF_8));
                // }

                // System.out.println("🎵 Added audio + timing to slide " + slideNum);

                slideNum++;
            }

            // Save modified PPTX
            // try (FileOutputStream fos = new FileOutputStream(outputPptx)) {
            // ppt.write(fos);
            // }

            // System.out.println("🎉 Completed. Saved: " + outputPptx);

            // -------------------------------------------
            // 👉 Print TOTAL CHARACTER COUNT
            // -------------------------------------------
            System.out.println("\n=================================================");
            System.out.println("🧮 TOTAL characters in all slides' speaker notes: " +
                    totalCharacters);
            System.out.println("=================================================\n");
        }
    }

    // ----------------------------------------------------------------------
    // 📌 Extract speaker notes
    // ----------------------------------------------------------------------
    private String extractNotes(XSLFSlide slide) {
        StringBuilder sb = new StringBuilder();

        XSLFNotes notes = slide.getNotes();

        if (notes != null) {
            for (XSLFShape shape : notes.getShapes()) {
                if (shape instanceof XSLFTextShape ts) {
                    sb.append(ts.getText()).append(" ");
                }
            }
        }

        return sb.toString().trim();
    }

    // ------------------------------------------------------------
    // 🧹 Clean speaker notes (remove unwanted symbols/emojis)
    // ------------------------------------------------------------

    private static String cleanNotes(String text) {
        if (text == null)
            return "";
        return removeEmojis(text);
    }

    // ⭐ Correct, modern emoji regex (covers 90%+ emojis)
    private static final String EMOJI_REGEX = "(?:[\\uD83C\\uDF00-\\uD83D\\uDDFF]|" +
            "[\\uD83E\\uDD00-\\uD83E\\uDDFF]|" +
            "[\\uD83D\\uDE00-\\uD83D\\uDE4F]|" +
            "[\\uD83C\\uDDE6-\\uD83C\\uDDFF]|" +
            "[\\u2600-\\u26FF]|" +
            "[\\u2700-\\u27BF]|" +
            "\\uFE0F|\\u20E3|" +
            // explicit keycap sequences: digit/#/* + optional VS16 + enclosing keycap
            "(?:[0-9#*]\\uFE0F?\\u20E3)" // ZWJ sequences (family emojis, profession emojis)
            + "(?:[\\uD83C-\\uDBFF\\uDC00-\\uDFFF]\\u200D[\\uD83C-\\uDBFF\\uDC00-\\uDFFF])+"
            + ")";

    public static boolean containsEmoji(String text) {
        return text.matches(".*(" + EMOJI_REGEX + ").*");
    }

    public static String removeEmojis(String text) {
        if (text == null)
            return "";

        // Remove emojis
        String cleaned = text.replaceAll(EMOJI_REGEX, "");

        // Remove zero-width joiner (ZWJ)
        cleaned = cleaned.replace("\u200D", "");

        // Fix multiple spaces
        cleaned = cleaned.replaceAll("\\s{2,}", " ").trim();

        // cleaned = cleaned.replaceAll("(?m)^(\\s*\\n){2,}", "\n").trim();

        return cleaned;
    }

    private int countCharacters(int slideNum, String cleanedText) {

        int count = cleanedText.length();
        totalCharacters += count;

        System.out.println("📝 Slide " + slideNum + " — Characters: " + count);
        return count;
    }
}
