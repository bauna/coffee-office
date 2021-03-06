package com.itcarg.office.aspose.word;

import java.io.File;
import java.io.InputStream;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;

import com.aspose.words.Cell;
import com.aspose.words.Document;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.MailMerge;
import com.aspose.words.MergeImageFieldEventArgs;
import com.aspose.words.MergeImageFieldEventHandler;
import com.aspose.words.Range;
import com.aspose.words.Run;
import com.aspose.words.VisitorAction;
import com.itcarg.office.aspose.BaseAspose;
import com.itcarg.office.word.WordHandler;

public class WordAspose extends BaseAspose implements WordHandler {

    private Document document;

    public WordAspose(InputStream docStream) throws Exception {
        document = new Document(docStream);
    }

    /**
     * @see com.itcarg.office.aspose.word.WordHandler#replaceMap(java.util.Map)
     */
    @Override
    public Map<String, Integer> replaceMap(Map<String, String> properties) throws Exception {
        HashMap<String, Integer> replaceCounts = new HashMap<String, Integer>() {
            private static final long serialVersionUID = 1L;

            @Override
            public Integer put(String key, Integer value) {
                Integer count = this.get(key);
                if (count == null) {
                    return super.put(key, value);
                }
                return super.put(key, value.intValue() + count.intValue());
            }
        };

        for (Map.Entry<String, String> entry : properties.entrySet()) {
            try {
                String replaceKey = "\\$" + entry.getKey();
                int count = getDocument().getRange().replace(Pattern.compile(replaceKey), entry.getValue());
                log.info("'{}' was replaced {} times", replaceKey, count);
                replaceCounts.put(entry.getKey(), count);
            } catch (Exception e) {
                log.error("error replacing text", e);
            }
        }

        return new HashMap<String, Integer>(replaceCounts);
    }

    public void mailMerge(final Map<String, Object> properties) throws Exception {
        
        MailMerge mailMerge = getDocument().getMailMerge();
        mailMerge.addMergeImageFieldEventHandler(new MergeImageFieldEventHandler() {
            
            @Override
            public void mergeImageField(Object sender, MergeImageFieldEventArgs e) throws Exception {
                URL url = (URL) e.getFieldValue();
                e.setImageStream(url.openStream());
            }
        });
        String fieldnames[] = new String[properties.size()];
        Object values[] = new Object[properties.size()];
        int i = 0;
        for (Map.Entry<String, Object> entry : properties.entrySet()) {
            fieldnames[i] = entry.getKey();
            values[i] = entry.getValue();
            i++;
        }
        mailMerge.execute(fieldnames, values);
    }
    
    /**
     * @see com.itcarg.office.word.WordHandler#saveAs(java.io.File)
     */
    @Override
    public void saveAs(File fileName) throws Exception {
        getDocument().save(fileName.getAbsolutePath());
    }

    private Document getDocument() {
        return document;
    }

    private class DocumentReplacer extends DocumentVisitor {
        private Map<String, String> properties;
        private HashMap<String, Integer> replaceCounts;

        public DocumentReplacer(Map<String, String> properties,
                HashMap<String, Integer> replaceCounts) {
            this.properties = properties;
            this.replaceCounts = replaceCounts;
        }
        
        private void replaceInRange(Range range) {
            for (Map.Entry<String, String> entry : properties.entrySet()) {
                try {
                    String replaceKey = "\\$" + entry.getKey();
                    int count = range.replace(Pattern.compile(replaceKey), entry.getValue());
                    log.info("'{}' was replaced {} times", replaceKey, count);
                    replaceCounts.put(entry.getKey(), count);
                } catch (Exception e) {
                    log.error("error replacing text", e);
                }
            }
        }

        @Override
        public int visitCellStart(Cell cell) throws Exception {
            log.info("replacing text in cell node");
            replaceInRange(cell.getRange());
            return VisitorAction.CONTINUE;
        }

        @Override
        public int visitRun(Run run) throws Exception {
            log.info("replacing text in run node");
            replaceInRange(run.getRange());
            return VisitorAction.CONTINUE;
        }
    }
}
