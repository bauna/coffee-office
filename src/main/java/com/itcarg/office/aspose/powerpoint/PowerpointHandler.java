package com.itcarg.office.aspose.powerpoint;

import java.io.File;
import java.io.IOException;
import java.util.Map;

public interface PowerpointHandler {

    void replaceMap(Map<String, String> properties) throws Exception;

    void saveAs(File filename) throws IOException;

}