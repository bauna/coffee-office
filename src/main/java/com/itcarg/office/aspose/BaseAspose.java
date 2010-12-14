package com.itcarg.office.aspose;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.aspose.words.License;

public class BaseAspose {
    protected final Logger log = LoggerFactory.getLogger(getClass());
    
    static {
        try {
            License license = new License();
            // Set the license of Aspose to avoid the evaluation
            // limitations
            license.setLicense(BaseAspose.class.getResourceAsStream("/aspose/Aspose.Total.Java.lic"));
        } catch (Exception e) {
            LoggerFactory.getLogger(BaseAspose.class).error("error setting the license", e);
        }
    }
}
