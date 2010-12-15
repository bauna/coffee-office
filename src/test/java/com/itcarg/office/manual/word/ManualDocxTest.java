package com.itcarg.office.manual.word;

import static org.testng.AssertJUnit.assertNotNull;
import static org.testng.AssertJUnit.assertTrue;

import java.io.File;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;

import com.itcarg.office.word.WordHandler;

public class ManualDocxTest {
    private AtomicInteger idGenerator = new AtomicInteger(0);
    
    private static final Logger log = LoggerFactory.getLogger(ManualDocxTest.class);

    @Test(threadPoolSize = 4, invocationCount = 20, timeOut = 10000,
            testName="docx-manual")
    public void testReplaceMapDocx() throws Exception {
        runFile("/word/test_document.docx", "ManualTest" + idGenerator.addAndGet(1) + ".docx");
    }

    private void runFile(String fileName, String out) throws Exception {
        log.info("start test for {}", fileName);
        WordHandler doc = new DocxImplementation(getClass().getResourceAsStream(fileName));
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

//        for (Map.Entry<String, Integer> entry : replaceMap.entrySet()) {
//            log.info("checking replaces for: {}", entry);
//            AssertJUnit.assertTrue(props.containsKey(entry.getKey()));
//            if ("key1".equals(entry.getKey())) {
//                assertEquals(2, entry.getValue().intValue());
//            } else if ("key2".equals(entry.getKey())) {
//                assertEquals(2, entry.getValue().intValue());
//            } else {
//                assertEquals(1, entry.getValue().intValue());
//            }
//        }
        log.info("end test for {}", fileName);
    }
}
