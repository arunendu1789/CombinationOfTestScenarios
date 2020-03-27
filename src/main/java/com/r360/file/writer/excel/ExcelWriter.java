package com.r360.file.writer.excel;

import com.r360.file.reader.excel.ExcelReader;
import com.r360.file.utils.FileUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class ExcelWriter {
    private static final Logger LOGGER = Logger.getLogger(ExcelWriter.class.getName());

    public void writeCombinationOfScenariosInExcelFile(Map<String, List<String>> mapOfTestCaseCombinations, String pathOfExcelFile) {
        Path excelFilePath = Paths.get(pathOfExcelFile);
        if(!excelFilePath.toFile().exists()) {
            try {
                Files.createFile(excelFilePath);
            } catch (IOException e) {
                LOGGER.error("Error while creating Excel File : "+e.getStackTrace().toString());
            }
        }
        if(excelFilePath.toFile().exists()) {
            String excelFileName = excelFilePath.getFileName().toString();
            LOGGER.info("FileName of the Excel Sheet : "+excelFileName);
            Workbook workBookOfExcelFile = FileUtils.getWorkBookOfExcelFile(excelFileName, excelFilePath.toFile());
            //Create New Excel Sheet
            Sheet fileSheet = workBookOfExcelFile.createSheet("TestCaseCombinations");
            fileSheet.setDefaultColumnWidth(30);

            //Create Style for Header Cells
            CellStyle cellStyle = workBookOfExcelFile.createCellStyle();
            Font excelFileFont = workBookOfExcelFile.createFont();
            excelFileFont.setFontName("Cambria");
            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
            excelFileFont.setBold(true);
            excelFileFont.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
            cellStyle.setFont(excelFileFont);

            //Write data in Excel File
            int rowCount = 0;
            Set<String> keySet = mapOfTestCaseCombinations.keySet();
            for(String key : keySet) {
                List<String> testCases = mapOfTestCaseCombinations.get(key);
                Row row = fileSheet.createRow(rowCount++);
                for(int i=0; i<testCases.size(); i++) {
                    row.createCell(i).setCellValue(testCases.get(i));
                }
            }
            writeScenariosInExcelFile(workBookOfExcelFile, excelFilePath);

        }
    }

    public void writeScenariosInExcelFile(Workbook workBookOfExcelFile, Path excelFilePath) {
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(excelFilePath.toFile());
            CellStyle cellStyle = workBookOfExcelFile.createCellStyle();
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setFillBackgroundColor((short) 245);
            workBookOfExcelFile.write(fileOutputStream);
        } catch (Exception e) {
            LOGGER.error("Error while writing data in Excel File : "+e.getStackTrace().toString());
        } finally {
            if(fileOutputStream!=null) {
                try {
                    fileOutputStream.close();
                } catch (IOException e) {
                    LOGGER.error("Error while closing FileOutputStream : "+e.getStackTrace().toString());
                }
            }
        }
    }

}
