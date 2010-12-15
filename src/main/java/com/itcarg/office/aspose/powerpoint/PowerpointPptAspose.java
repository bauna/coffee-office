package com.itcarg.office.aspose.powerpoint;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

import com.aspose.slides.Placeholder;
import com.aspose.slides.Placeholders;
import com.aspose.slides.Presentation;
import com.aspose.slides.Slide;
import com.aspose.slides.Slides;
import com.aspose.slides.TextHolder;
import com.itcarg.office.aspose.BaseAspose;

public class PowerpointAspose extends BaseAspose {
    private Presentation presentation;

    public PowerpointAspose(InputStream slidesStream) {
        presentation = new Presentation(slidesStream);
    }

    private Presentation getPresentation() {
        return presentation;
    }

    public void replaceMap(Map<String, String> properties) throws Exception {
        Slides slides = getPresentation().getSlides();
        for (int i = 0; i < slides.size(); i++) {
            Slide slide = slides.get(i);
            Placeholders placeholders = slide.getPlaceholders();
            for (int j = 0; j < placeholders.size(); j++) {
                Placeholder placeholder = placeholders.get(j);
                if (placeholder instanceof TextHolder) {
                    TextHolder textHolder = (TextHolder) placeholder;
                    for (Map.Entry<String, String> entry : properties.entrySet()) {
                        textHolder.setText(textHolder.getText().replaceAll("\\$" + entry.getKey(),
                                entry.getValue()));
                    }
                }
            }
        }
    }

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
