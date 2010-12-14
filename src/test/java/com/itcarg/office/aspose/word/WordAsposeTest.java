package com.itcarg.office.aspose.word;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.util.LinkedHashMap;
import java.util.Map;

import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class WordAsposeTest {
    private static final Logger log = LoggerFactory.getLogger(WordAsposeTest.class);
    
    @Test
    public void testReplaceMapDocx() throws Exception {
        testFile("/word/test_document.docx", "test.docx");
    }
    
    @Test
    public void testReplaceMapDoc() throws Exception {
        testFile("/word/test_document.doc", "test.doc");
    }
    
    private void testFile(String fileName, String out) throws Exception {
        log.info("start test for {}", fileName);
        WordAspose doc = new WordAspose(getClass().getResourceAsStream(fileName));
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
            assertTrue(props.containsKey(entry.getKey()));
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
}
