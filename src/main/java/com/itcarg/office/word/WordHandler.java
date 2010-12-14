package com.itcarg.office.word;

import java.io.File;
import java.util.Map;

public interface WordHandler {

    Map<String, Integer> replaceMap(Map<String, String> properties) throws Exception;

    void saveAs(File fileName) throws Exception;

}