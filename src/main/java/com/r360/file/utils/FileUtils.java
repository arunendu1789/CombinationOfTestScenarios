package com.r360.file.utils;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class FileUtils {
    private static final Logger LOGGER = Logger.getLogger(FileUtils.class.getName());

    public static Workbook getWorkBookOfExcelFile(String fileName, File file) {
        FileInputStream fileInputStream = null;
        try {
            if(fileName == null) {
                return null;
            } else if(fileName.endsWith("xlsx")) {
                if(file.length()>0) {
                    fileInputStream = new FileInputStream(file);
                    return new XSSFWorkbook(fileInputStream);
                } else {
                    return new XSSFWorkbook();
                }
            } else if(fileName.endsWith("xls")) {
                if(file.length()>0) {
                    fileInputStream = new FileInputStream(file);
                    return new HSSFWorkbook(fileInputStream);
                } else {
                    return new HSSFWorkbook();
                }
            }
        } catch (Exception e) {
            LOGGER.error("An error occurred while getting Workbook object : "+ e.getStackTrace().toString());
        } finally {
            if(fileInputStream!=null) {
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    LOGGER.error("An error occurred while getting Workbook object : "+ e.getStackTrace().toString());
                }
            }
        }
        return null;
    }
}
