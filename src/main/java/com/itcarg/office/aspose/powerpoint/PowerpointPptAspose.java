package com.itcarg.office.aspose.powerpoint;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

import com.aspose.slides.Paragraph;
import com.aspose.slides.Paragraphs;
import com.aspose.slides.Placeholder;
import com.aspose.slides.Placeholders;
import com.aspose.slides.Presentation;
import com.aspose.slides.Slide;
import com.aspose.slides.Slides;
import com.aspose.slides.TextHolder;
import com.itcarg.office.aspose.BaseAspose;

public class PowerpointPptAspose extends BaseAspose implements PowerpointHandler {
    private Presentation presentation;

    public PowerpointPptAspose(InputStream slidesStream) {
        presentation = new Presentation(slidesStream);
    }

    private Presentation getPresentation() {
        return presentation;
    }

    /* (non-Javadoc)
     * @see com.itcarg.office.aspose.powerpoint.PowerpointHandler#replaceMap(java.util.Map)
     */
    @Override
    public void replaceMap(Map<String, String> properties) throws Exception {
        Slides slides = getPresentation().getSlides();
        for (int i = 0; i < slides.size(); i++) {
            Slide slide = slides.get(i);
            Placeholders placeholders = slide.getPlaceholders();
            for (int j = 0; j < placeholders.size(); j++) {
                Placeholder placeholder = placeholders.get(j);
                if (placeholder instanceof TextHolder) {
                    TextHolder textHolder = (TextHolder) placeholder;
                    Paragraphs paragraphs = textHolder.getParagraphs();
                    for (int k = 0; k < paragraphs.size(); k++) {
                        Paragraph paragraph = paragraphs.get(k);
                        for (Map.Entry<String, String> entry : properties.entrySet()) {
                            paragraph.setText(paragraph.getText()
                                    .replaceAll("\\$" + entry.getKey(),
                                            entry.getValue()));
                        }
                    }
                }
            }
        }
    }

    /* (non-Javadoc)
     * @see com.itcarg.office.aspose.powerpoint.PowerpointHandler#saveAs(java.io.File)
     */
    @Override
    public void saveAs(File filename) throws IOException {
        OutputStream out = null;
        try {
            out = new BufferedOutputStream(new FileOutputStream(filename));
            getPresentation().write(out);
        } finally {
            if (out != null) {
                out.close();
            }
        }
    }
}
