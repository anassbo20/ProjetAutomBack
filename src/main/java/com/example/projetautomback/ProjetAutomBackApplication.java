package com.example.projetautomback;

import org.apache.poi.xslf.usermodel.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

@SpringBootApplication
public class ProjetAutomBackApplication {

    public static void main(String[] args) throws IOException {
        SpringApplication.run(ProjetAutomBackApplication.class, args);

        FileInputStream fis = new FileInputStream("src/main/resources/template.pptx");
        XMLSlideShow ppt = new XMLSlideShow(fis);
        List<XSLFSlide> slidesList = ppt.getSlides();
        XSLFSlide[] slidesArray = slidesList.toArray(new XSLFSlide[slidesList.size()]);

        String searchText = "Rapport";
        for (XSLFSlide slide : slidesArray) {
            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) shape;
                    for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                        for (XSLFTextRun run : paragraph.getTextRuns()) {
                            String runText = run.getRawText();
                            if (runText.contains(searchText)) {
                                String replacementText = runText.replaceAll(searchText, "test");
                                run.setText(replacementText);
                            }
                        }
                    }
                }
            }
        }

        FileOutputStream outputStream = new FileOutputStream("src/main/resources/modified_template.pptx");
        ppt.write(outputStream);
        fis.close();
        outputStream.close();
    }
}
