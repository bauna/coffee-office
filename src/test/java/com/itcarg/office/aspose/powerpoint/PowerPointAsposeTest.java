package com.itcarg.office.aspose.powerpoint;

import static org.testng.AssertJUnit.assertTrue;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;

public class PowerPointAsposeTest {
    private AtomicInteger idGenerator = new AtomicInteger(0);

    private static final Logger log = LoggerFactory.getLogger(PowerPointAsposeTest.class);

    @Test(threadPoolSize = 4, invocationCount = 20, timeOut = 10000)
    public void testReplaceMapPptx() throws Exception {
        log.info("loading powerpoint file: {}", "/powerp/test.pptx");
        runFile(new PowerpointPptxAspose(getFile("/powerp/test.pptx")),
                "test" + idGenerator.addAndGet(1) + ".pptx");
        log.info("end test for {}", "/powerp/test.pptx");
    }

    @Test(threadPoolSize = 4, invocationCount = 20, timeOut = 10000)
    public void testReplaceMapPpt() throws Exception {
        log.info("loading powerpoint file: {}", "/powerp/test.ppt");
        runFile(new PowerpointPptAspose(getFile("/powerp/test.ppt")),
                "test" + idGenerator.addAndGet(1) + ".ppt");
        log.info("end test for {}", "/powerp/test.ppt");
    }

    private InputStream getFile(String filename) {
        return new BufferedInputStream(getClass().getResourceAsStream(filename));
    }
    
    private void runFile(PowerpointHandler doc, String out) throws Exception {
        Map<String, String> props = new LinkedHashMap<String, String>();

        for (int i = 1; i < 8; i++) {
            props.put("key" + i, "value" + i);
        }
        doc.replaceMap(props);

        File fileOut = new File(System.getProperty("java.io.tmpdir"), out);
        fileOut.delete();
        doc.saveAs(fileOut);
        assertTrue(fileOut.exists());
    }
}
