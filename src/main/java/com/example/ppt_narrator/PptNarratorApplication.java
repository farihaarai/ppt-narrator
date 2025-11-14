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

import com.mpatric.mp3agic.Mp3File;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;

@SpringBootApplication
public class PptNarratorApplication implements CommandLineRunner {

	// AWS Polly client (used to generate speech audio)
	@Autowired
	private PollyClient polly;

	public static void main(String[] args) {
		SpringApplication.run(PptNarratorApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {

		// Check if a PPTX file path is provided in arguments
		if (args.length == 0) {
			System.err.println(
					" Usage: mvn spring-boot:run -Dspring-boot.run.arguments=\"<path-to-pptx> [output-directory]\"");
			System.exit(1);
		}

		String pptxFile = args[0];
		String outputDir = args.length > 1 ? args[1] : "slide_audios";

		System.out.println(" Input PPTX: " + pptxFile);
		System.out.println(" Output Directory: " + outputDir);

		// Create folder for saving mp3 files
		Files.createDirectories(Paths.get(outputDir));

		// Create a copy of the original PPTX because we will modify it
		String outputPptx = pptxFile.replace(".pptx", "_with_audio.pptx");
		Files.copy(Paths.get(pptxFile), Paths.get(outputPptx), StandardCopyOption.REPLACE_EXISTING);

		// Open the copied PPTX
		try (FileInputStream fis = new FileInputStream(outputPptx);
				XMLSlideShow ppt = new XMLSlideShow(OPCPackage.open(fis))) {

			int slideNum = 1;

			// Loop through each slide in the presentation
			for (XSLFSlide slide : ppt.getSlides()) {

				// Extract the notes text from the slide (speaker notes)
				String notesText = extractNotes(slide);

				// If slide has no notes, skip audio creation
				if (notesText.isBlank()) {
					System.out.println("Slide " + slideNum + ": No notes found, skipping audio generation.");
					slideNum++;
					continue;
				}

				// Filename for the audio for this slide
				String audioFileName = outputDir + "/slide_" + slideNum + ".mp3";
				Path audioPath = Paths.get(audioFileName);

				// If audio file does NOT exist, generate speech using AWS Polly
				if (!Files.exists(audioPath)) {
					SynthesizeSpeechRequest request = SynthesizeSpeechRequest.builder()
							.text(notesText)
							.voiceId("Danielle") // Voice used by Polly
							.outputFormat(OutputFormat.MP3)
							.build();

					// Create MP3 file
					try (ResponseInputStream<SynthesizeSpeechResponse> speechStream = polly.synthesizeSpeech(request);
							OutputStream out = new FileOutputStream(audioFileName)) {
						speechStream.transferTo(out);
					}
					System.out.println(" Generated audio: " + audioFileName);
				} else {
					System.out.println(" Using existing audio: " + audioFileName);
				}

				// Embed audio inside the PPTX file
				OPCPackage pkg = ppt.getPackage();
				String mediaPath = "/ppt/media/audio" + slideNum + ".mp3";

				// Create a part name inside the ppt/media folder
				PackagePartName mediaPartName = PackagingURIHelper.createPartName(mediaPath);

				// If audio already exists in PPT, remove old one
				if (pkg.containPart(mediaPartName)) {
					pkg.removePart(mediaPartName);
				}

				// Create new audio part in PPTX
				PackagePart mediaPart = pkg.createPart(mediaPartName, "audio/mpeg");

				// Copy our MP3 file into the PPTX package
				try (InputStream is = Files.newInputStream(audioPath);
						OutputStream os = mediaPart.getOutputStream()) {
					is.transferTo(os);
				}

				// Create relationship between slide and audio part
				String relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/media";
				PackageRelationship rel = slide.getPackagePart().addRelationship(mediaPartName, TargetMode.INTERNAL,
						relType);
				String relId = rel.getId();

				System.out.println(" Embedded audio for slide " + slideNum + " as relationship id: " + relId);

				// Read MP3 duration to control slide auto-advance timing
				Mp3File mp3 = new Mp3File(audioFileName);
				int durationSeconds = (int) Math.ceil(mp3.getLengthInSeconds());
				System.out.println(" Duration for slide " + slideNum + ": " + durationSeconds + "s");

				// Modify slide XML to attach audio and transition timing
				PackagePart slidePart = slide.getPackagePart();
				String slideXml;

				// Read slide XML
				try (InputStream slideIs = slidePart.getInputStream()) {
					slideXml = new String(slideIs.readAllBytes(), StandardCharsets.UTF_8);
				}

				// Add audio XML snippet referencing the MP3 file
				String audioSnippet = "  <p:audio xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" "
						+
						"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
						"    <p:audioFile r:embed=\"" + relId + "\"/>" +
						"  </p:audio>\n";

				// Insert audio before closing spTree tag
				String marker = "</p:spTree>";
				if (slideXml.contains(marker)) {
					slideXml = slideXml.replace(marker, audioSnippet + marker);
				} else {
					// If structure different, append at end
					slideXml += "\n" + audioSnippet;
				}

				// Add/Modify slide transition timing (auto-advance after audio ends)
				String transitionOpenTag = "<p:transition";

				if (slideXml.contains(transitionOpenTag)) {
					// Add advTm attribute if transition tag exists
					slideXml = slideXml.replaceFirst("<p:transition([^>]*)>",
							"<p:transition$1 advTm=\"" + (durationSeconds * 1000) + "\">");
				} else {
					// If no transition tag, insert one inside <p:cSld>
					String csldMarker = "<p:cSld";
					int idx = slideXml.indexOf(csldMarker);

					if (idx != -1) {
						int endOpen = slideXml.indexOf('>', idx);
						if (endOpen != -1) {
							String toInsert = "\n  <p:transition advTm=\"" + (durationSeconds * 1000) + "\"/>\n";
							slideXml = slideXml.substring(0, endOpen + 1) + toInsert + slideXml.substring(endOpen + 1);
						}
					} else {
						// If everything fails, add transition at end
						slideXml += "\n  <p:transition advTm=\"" + (durationSeconds * 1000) + "\"/>\n";
					}
				}

				// Save updated slide XML back into PPTX
				try (OutputStream os = slidePart.getOutputStream()) {
					os.write(slideXml.getBytes(StandardCharsets.UTF_8));
				}

				System.out.println("🎵 Injected audio element + transition for slide " + slideNum);

				slideNum++;
			}

			// Save final PPTX with audio embedded
			try (FileOutputStream fos = new FileOutputStream(outputPptx)) {
				ppt.write(fos);
			}

			System.out.println(" All done! Saved: " + outputPptx);
		}
	}

	// Extracts the speaker notes text from a slide
	private String extractNotes(XSLFSlide slide) {
		StringBuilder sb = new StringBuilder();
		XSLFNotes notes = slide.getNotes();

		if (notes != null) {
			for (XSLFShape shape : notes.getShapes()) {
				if (shape instanceof XSLFTextShape textShape) {
					sb.append(textShape.getText()).append(" ");
				}
			}
		}

		return sb.toString().trim();
	}
}
