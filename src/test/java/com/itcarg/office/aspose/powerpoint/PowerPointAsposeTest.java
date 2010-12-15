package com.itcarg.office.aspose.powerpoint;

import static org.testng.AssertJUnit.assertNotNull;
import static org.testng.AssertJUnit.assertTrue;

import java.io.File;
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
        runFile("/powerp/test.pptx", "test" + idGenerator.addAndGet(1) + ".pptx");
    }

    @Test(threadPoolSize = 4, invocationCount = 20, timeOut = 10000)
    public void testReplaceMapPpt() throws Exception {
        runFile("/powerp/test.ppt", "test" + idGenerator.addAndGet(1) + ".ppt");
    }

    private void runFile(String fileName, String out) throws Exception {
        log.info("start test for {}", fileName);
        PowerpointAspose doc = new PowerpointAspose(getClass().getResourceAsStream(fileName));
        assertNotNull(doc);

        Map<String, String> props = new LinkedHashMap<String, String>();

        for (int i = 1; i < 8; i++) {
            props.put("key" + i, "value" + i);
        }
        doc.replaceMap(props);

        File fileOut = new File(System.getProperty("java.io.tmpdir"), out);
        fileOut.delete();
        doc.saveAs(fileOut);
        assertTrue(fileOut.exists());

        log.info("end test for {}", fileName);
    }
}
