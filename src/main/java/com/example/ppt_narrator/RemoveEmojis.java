// package com.example.ppt_narrator;

// import org.apache.poi.openxml4j.opc.OPCPackage;
// import org.apache.poi.xslf.usermodel.*;

// import org.springframework.boot.CommandLineRunner;
// import org.springframework.boot.SpringApplication;
// import org.springframework.boot.autoconfigure.SpringBootApplication;

// import java.io.FileInputStream;
// import java.io.FileOutputStream;
// import java.io.IOException;
// import java.nio.file.*;

// @SpringBootApplication
// public class PptNarratorApplication implements CommandLineRunner {

// // 👉 Count total characters across all slides
// private int totalCharacters = 0;

// public static void main(String[] args) {
// SpringApplication.run(PptNarratorApplication.class, args);
// }

// @Override
// public void run(String... args) throws Exception {

// if (args.length == 0) {
// System.err.println(
// "❌ Usage: mvn spring-boot:run -Dspring-boot.run.arguments=\"<path-to-pptx>
// [output-directory]\"");
// System.exit(1);
// }

// String pptxFile = args[0];
// String outputDir = args.length > 1 ? args[1] : "slide_audios";

// System.out.println("📂 Input PPTX: " + pptxFile);
// System.out.println("📁 Output Directory: " + outputDir);

// // Create output directory if missing
// Files.createDirectories(Paths.get(outputDir));

// // ------------------------------
// // 📄 Create processed copy of PPTX
// // ------------------------------
// Path inputPath = Paths.get(pptxFile);
// String originalFileName = inputPath.getFileName().toString();
// String processedFileName = originalFileName.replace(".pptx",
// "_processed.pptx");

// Path processedPptxPath = Paths.get(outputDir, processedFileName);
// Files.copy(inputPath, processedPptxPath,
// StandardCopyOption.REPLACE_EXISTING);

// System.out.println("✅ Processed PPTX created at: " +
// processedPptxPath.toAbsolutePath());

// // ------------------------------
// // 📘 Open and process the copied PPTX
// // ------------------------------
// try (FileInputStream fis = new FileInputStream(processedPptxPath.toFile());
// XMLSlideShow ppt = new XMLSlideShow(OPCPackage.open(fis))) {

// int slideNum = 1;

// for (XSLFSlide slide : ppt.getSlides()) {

// String notesText = extractNotes(slide);
// String cleanedNotes = cleanNotes(notesText);

// // ✔ Write cleaned text back into PPT notes
// XSLFNotes notes = slide.getNotes();
// if (notes != null) {
// for (XSLFShape shape : notes.getShapes()) {
// if (shape instanceof XSLFTextShape ts) {
// ts.clearText();
// ts.setText(cleanedNotes); // ✔ override notes text
// }
// }
// }

// // ✔ Count characters AFTER cleaning
// countCharacters(slideNum, cleanedNotes);
// slideNum++;
// }

// // ✔ IMPORTANT: Save the updated PPTX
// try (FileOutputStream out = new FileOutputStream(processedPptxPath.toFile()))
// {
// ppt.write(out);
// }

// } catch (IOException e) {
// throw new RuntimeException("❌ Failed to open processed PPTX: " +
// e.getMessage(), e);
// }

// System.out.println("🏁 Total Characters Across All Slides: " +
// totalCharacters);
// }

// // ------------------------------------------------------------
// // 📝 Extract Notes from Slide
// // ------------------------------------------------------------
// private String extractNotes(XSLFSlide slide) {
// StringBuilder sb = new StringBuilder();

// XSLFNotes notes = slide.getNotes();
// if (notes != null) {
// for (XSLFShape shape : notes.getShapes()) {
// if (shape instanceof XSLFTextShape ts) {
// sb.append(ts.getText()).append(" ");
// }
// }
// }
// return sb.toString().trim();
// }

// // ------------------------------------------------------------
// // 🧹 Clean notes (remove ALL emojis)
// // ------------------------------------------------------------
// private static String cleanNotes(String text) {
// if (text == null)
// return "";
// return removeEmojis(text);
// }

// // ⭐ Correct, modern emoji regex (covers 90%+ emojis)
// private static final String EMOJI_REGEX =
// "(?:[\\uD83C\\uDF00-\\uD83D\\uDDFF]|" +
// "[\\uD83E\\uDD00-\\uD83E\\uDDFF]|" +
// "[\\uD83D\\uDE00-\\uD83D\\uDE4F]|" +
// "[\\uD83C\\uDDE6-\\uD83C\\uDDFF]|" +
// "[\\u2600-\\u26FF]|" +
// "[\\u2700-\\u27BF]|" +
// "\\uFE0F|\\u20E3|" +
// // explicit keycap sequences: digit/#/* + optional VS16 + enclosing keycap
// "(?:[0-9#*]\\uFE0F?\\u20E3)" +
// ")";

// public static boolean containsEmoji(String text) {
// return text.matches(".*(" + EMOJI_REGEX + ").*");
// }

// public static String removeEmojis(String text) {
// return text.replaceAll(EMOJI_REGEX, "");
// }

// // ------------------------------------------------------------
// // 🔢 Count characters
// // ------------------------------------------------------------
// private void countCharacters(int slideNum, String cleanedText) {
// int count = cleanedText.length();
// totalCharacters += count;

// System.out.println("📝 Slide " + slideNum + " — Characters: " + count);
// }
// }
