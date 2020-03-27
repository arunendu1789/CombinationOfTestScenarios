package com.r360.test.file.excel;

import com.r360.file.reader.excel.ExcelReader;
import com.r360.file.writer.excel.ExcelWriter;

import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

public class TestCreationOfCombinationInExcelFile {
    public static void main(String[] args) {
        ExcelReader excelReader = new ExcelReader();
        ExcelWriter excelWriter = new ExcelWriter();

        String excelFilePath = Paths.get("data").toString()+"/preTestCases.xlsx";
        Map<String, List<String>> mappedCombinationOfTestCases =
                excelReader.getMappedCombinationOfTestCases(excelFilePath);
        excelFilePath = Paths.get("data").toString()+"/preTestCasesOutput.xlsx";
        excelWriter.writeCombinationOfScenariosInExcelFile(mappedCombinationOfTestCases, excelFilePath);
    }
}
