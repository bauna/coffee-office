package com.itcarg.office.manual.word;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.itcarg.office.word.WordHandler;

public class DocxImplementation implements WordHandler {
    private static final Logger log = LoggerFactory.getLogger(DocxImplementation.class);
    private byte[] document;

    public DocxImplementation(InputStream docStream) throws IOException {
        BufferedInputStream in = new BufferedInputStream(docStream);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        int count;
        byte[] buf = new byte[2000];
        while ((count = in.read(buf)) != -1) {
            out.write(buf, 0, count);
        }
        document = out.toByteArray();
    }

    @Override
    public Map<String, Integer> replaceMap(Map<String, String> properties) throws Exception {
        ZipInputStream in = null;
        ZipOutputStream out = null;
        ByteArrayOutputStream newDoc = new ByteArrayOutputStream(document.length);
        try {
            in = new ZipInputStream(new ByteArrayInputStream(document));
            out = new ZipOutputStream(newDoc);
            ZipEntry entry;
            byte[] data = new byte[2000];
            int count;
            while ((entry = in.getNextEntry()) != null) {
                out.putNextEntry(new ZipEntry(entry));
                if (!entry.isDirectory() && entry.getName().startsWith("word/") 
                        && entry.getName().endsWith(".xml")) {
                    log.info("begin replacing properties", data);
                    ByteArrayOutputStream docData = new ByteArrayOutputStream();
                    while ((count = in.read(data)) != -1) {
                        docData.write(data, 0, count);
                    }
                    String content = new String(docData.toByteArray());
                    for (Map.Entry<String, String> prop : properties.entrySet()) {
                        String regexKey = "\\$" + prop.getKey();
                        log.info("replacing '{}' by '{}'", regexKey, prop.getValue());
                        content = content.replaceAll(regexKey, prop.getValue());
                    }
                    out.write(content.getBytes());
                } else {
                    log.info("passthrough entry: {}", entry.getName());
                    while ((count = in.read(data)) != -1) {
                        out.write(data, 0, count);
                    }
                }
                out.closeEntry();
            }
        } finally {
            if (in != null) {
                in.close();
            }
            if (out != null) {
                out.close();
                document = newDoc.toByteArray();
            }
        }
        return null;
    }

    @Override
    public void saveAs(File fileName) throws Exception {
        OutputStream out = null;
        try {
            out = new BufferedOutputStream(new FileOutputStream(fileName));
            out.write(document);
        } finally {
            if (out != null) {
                out.close();
            }
        }

    }

}
