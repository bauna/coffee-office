package com.itcarg.office.aspose.powerpoint;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

import com.aspose.slides.pptx.AutoShapeEx;
import com.aspose.slides.pptx.PresentationEx;
import com.aspose.slides.pptx.ShapeEx;
import com.aspose.slides.pptx.ShapesEx;
import com.aspose.slides.pptx.SlideEx;
import com.aspose.slides.pptx.SlidesEx;
import com.aspose.slides.pptx.TextFrameEx;
import com.itcarg.office.aspose.BaseAspose;

public class PowerpointPptxAspose extends BaseAspose implements PowerpointHandler {
    private PresentationEx presentation;

    public PowerpointPptxAspose(InputStream slidesStream) throws IOException {
        presentation = new PresentationEx(slidesStream);
    }

    private PresentationEx getPresentation() {
        return presentation;
    }

    public void replaceMap(Map<String, String> properties) throws Exception {
        SlidesEx slides = getPresentation().getSlides();
        for (int i = 0; i < slides.size(); i++) {
            SlideEx slide = slides.get(i);
            ShapesEx shapes = slide.getShapes();
            for (int j = 0; j < shapes.size(); j++) {
                ShapeEx shape = shapes.get(j);
                if (shape instanceof AutoShapeEx) {
                    AutoShapeEx autoShape = (AutoShapeEx) shape;
                    TextFrameEx textFrame = autoShape.getTextFrame();
                    for (Map.Entry<String, String> entry : properties.entrySet()) {
                        textFrame.setText(textFrame.getText()
                                .replaceAll("\\$" + entry.getKey(),
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
