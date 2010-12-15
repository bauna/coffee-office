package com.itcarg.office.aspose.word;

import static org.testng.AssertJUnit.assertEquals;
import static org.testng.AssertJUnit.assertNotNull;
import static org.testng.AssertJUnit.assertTrue;

import java.io.BufferedInputStream;
import java.io.File;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;

import com.itcarg.office.word.WordHandler;

public class WordAsposeTest {
    private AtomicInteger idGenerator = new AtomicInteger(0);
    
    private static final Logger log = LoggerFactory.getLogger(WordAsposeTest.class);

    @Test(threadPoolSize = 4, invocationCount = 20, timeOut = 10000,
            testName="docx-aspose")
    public void testReplaceMapDocx() throws Exception {
        runFile("/word/test_document.docx", "test" + idGenerator.addAndGet(1) + ".docx");
    }

    @Test(threadPoolSize = 4, invocationCount = 20, timeOut = 10000,
            testName="doc-aspose")
    public void testReplaceMapDoc() throws Exception {
        runFile("/word/test_document.doc", "test" + idGenerator.addAndGet(1) + ".doc");
    }

    private void runFile(String fileName, String out) throws Exception {
        log.info("start test for {}", fileName);
        WordHandler doc = new WordAspose(new BufferedInputStream(
                getClass().getResourceAsStream(fileName)));
        assertNotNull(doc);

        Map<String, String> props = new LinkedHashMap<String, String>();

        for (int i = 1; i < 8; i++) {
            props.put("key" + i, "value" + i);
        }
        Map<String, Integer> replaceMap = doc.replaceMap(props);

        File fileOut = new File(System.getProperty("java.io.tmpdir"), out);
        fileOut.delete();
        doc.saveAs(fileOut);
        assertTrue(fileOut.exists());

        for (Map.Entry<String, Integer> entry : replaceMap.entrySet()) {
            log.info("checking replaces for: {}", entry);
            AssertJUnit.assertTrue(props.containsKey(entry.getKey()));
            if ("key1".equals(entry.getKey())) {
                assertEquals(2, entry.getValue().intValue());
            } else if ("key2".equals(entry.getKey())) {
                assertEquals(2, entry.getValue().intValue());
            } else {
                assertEquals(1, entry.getValue().intValue());
            }
        }
        log.info("end test for {}", fileName);
    }
    
    @Test(threadPoolSize = 4, invocationCount = 20, timeOut = 10000,
            testName="doc-images")
    public void testMailMerge() throws Exception {
       Map<String,Object> props = new HashMap<String, Object>();
       Map<String, String> replaces = new HashMap<String, String>();
       WordAspose doc = new WordAspose(new BufferedInputStream(
               getClass().getResourceAsStream("/word/image_mailmerge.doc")));
       replaces.put("key1", "value1");
       props.put("name", "Jane Doe");
       props.put("image1", getClass().getResource("/word/test.jpg"));
       
       doc.replaceMap(replaces);
       doc.mailMerge(props);
       
       doc.saveAs(new File(System.getProperty("java.io.tmpdir"), 
               "mailMerge" + idGenerator.addAndGet(1) + ".doc"));
    }
}
