package com.r360.file.reader.excel;

import com.r360.file.utils.FileUtils;
import org.apache.commons.compress.utils.Lists;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.*;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.IntStream;

public class ExcelReader {
    private static final Logger LOGGER = Logger.getLogger(ExcelReader.class.getName());

    private static int rowLength;

    public Map<String, List<String>> getMappedCombinationOfTestCases(String pathOfExcelFile) {
        Map<String, List<String>> mapOfTestCaseCombinations = new LinkedHashMap<String, List<String>>();

        Path pathOfExcelSheet = Paths.get(pathOfExcelFile);
        if(pathOfExcelSheet.toFile().exists()) {
            String excelFileName = pathOfExcelSheet.getFileName().toString();
            LOGGER.info("FileName of the Excel Sheet : "+excelFileName);
            Workbook workbook = FileUtils.getWorkBookOfExcelFile(excelFileName, pathOfExcelSheet.toFile());
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = firstSheet.iterator();
            List<Row> rowList = Lists.newArrayList(rowIterator);
            if(rowList.size()>0) {
                rowList.stream().forEach(row -> {
                    Iterator<Cell> cellIterator = row.cellIterator();
                    List<Cell> cellList = Lists.newArrayList(cellIterator);
                    if(cellList.size()>0) {
                        boolean isFirstIndex = row.equals(rowList.get(0));
                        List<String> stringObjectOfTestCase = new LinkedList<String>();
                        List<String[]> testScenarios = new LinkedList<>();
                        rowLength = cellList.size();
                        IntStream.range(0, cellList.size()).forEachOrdered(index ->  {
                            Cell cell = cellList.get(index);
                            String headerString = "TestScenarioHeaders";
                            int column = cell.getAddress().getColumn();
                            LOGGER.info("Cell Number : "+ column);
                            switch (column) {
                                case 0 :
                                    if(cell==cellList.get(0)) {
                                        headerString = cell.getStringCellValue();
                                    }
                                    stringObjectOfTestCase.add(cell.getStringCellValue());
                                    break;
                                case 1 :
                                case 2 :
                                case 3 :
                                case 4 :
                                    stringObjectOfTestCase.add(cell.getStringCellValue());
                                    break;
                                case 5 :
                                case 6 :
                                case 7 :
                                case 8 :
                                case 9 :
                                case 10 :
                                case 11 :
                                case 12 :
                                case 13 :
                                case 14 :
                                case 15 :
                                case 16 :
                                case 17 :
                                case 18 :
                                case 19 :
                                case 20 :
                                case 21 :
                                    cell.setCellType(CellType.STRING);
                                    String cellValue = cell.getStringCellValue();
                                    String[] testScenario;
                                    if(cellValue.contains(",")) {
                                        testScenario = cellValue.split(",");
                                    } else {
                                        testScenario = new String[]{cellValue};
                                    }
                                    testScenarios.add(testScenario);
                                    break;
                                }
                                if(column==22 && isFirstIndex) {
                                    String cellValue = cell.getStringCellValue();
                                    String[] testScenario = new String[]{cellValue};
                                    testScenarios.add(testScenario);
                                }
                            });
                        List<List<String>> combinationOfScenarios = new LinkedList<>();
                        combinationOfScenarios = getCombinationOfScenarios(testScenarios, stringObjectOfTestCase, combinationOfScenarios);
                        int index = 0;
                        for(List<String> combinationOfScenario : combinationOfScenarios) {
                            mapOfTestCaseCombinations.put(stringObjectOfTestCase.get(0)+"-"+(index++), combinationOfScenario);
                        }
                    }
                });
                return mapOfTestCaseCombinations;
            }
        }
        return null;
    }

    public List<List<String>> getCombinationOfScenarios(List<String[]> testScenarios, List<String> stringObjectOfTestCase, List<List<String>> listOfTestCaseObjects) {
        for(int i=0; i<testScenarios.size(); i++) {
            String[] testScenarioCases = testScenarios.get(i);
            if(testScenarioCases.length > 1) {
                for(int j=0; j<testScenarioCases.length; j++) {
                    stringObjectOfTestCase.add(testScenarioCases[j]);
                    List<String[]> childTestScenarios = new LinkedList<>();
                    for(int k=i+1; k<testScenarios.size(); k++) {
                        childTestScenarios.add(testScenarios.get(k));
                    }
                    if(childTestScenarios.size()>0) {
                        listOfTestCaseObjects = getCombinationOfScenarios(childTestScenarios, stringObjectOfTestCase, listOfTestCaseObjects);
                    } else {
                        List<String> stringObjectOfTestCaseCopy = new LinkedList<>(stringObjectOfTestCase);
                        setAmountBasedOnBaseAmount(stringObjectOfTestCaseCopy);
                        listOfTestCaseObjects.add(stringObjectOfTestCaseCopy);
                    }
                    if(stringObjectOfTestCase.get(stringObjectOfTestCase.size()-1).equalsIgnoreCase(testScenarioCases[j])) {
                        stringObjectOfTestCase.remove(stringObjectOfTestCase.size()-1);
                    } else {
                        int fromIndex = stringObjectOfTestCase.indexOf(testScenarioCases[j]);
                        for (int l = stringObjectOfTestCase.size() - 1; l >= fromIndex; l--) {
                            stringObjectOfTestCase.remove(l);
                        }
                    }
                }
                break;
            } else {
                stringObjectOfTestCase.add(testScenarioCases[0]);
            }
        }
        if(stringObjectOfTestCase.size()==(rowLength-1)) {
            List<String> stringObjectOfTestCaseCopy = new LinkedList<>(stringObjectOfTestCase);
            setAmountBasedOnBaseAmount(stringObjectOfTestCaseCopy);
            listOfTestCaseObjects.add(stringObjectOfTestCaseCopy);
        } else if(stringObjectOfTestCase.size()==rowLength) {
            List<String> stringObjectOfTestCaseCopy = new LinkedList<>(stringObjectOfTestCase);
            listOfTestCaseObjects.add(stringObjectOfTestCaseCopy);
        }
        return listOfTestCaseObjects;
    }


    public static void main(String[] args) {
        List<String[]> testScenarios = new LinkedList<>();
        String[] strings1 = new String[]{"null"};
        String[] strings2 = new String[]{"0", "1"};
        String[] strings3 = new String[]{"null"};
        String[] strings4 = new String[]{"8","9"};
        String[] strings5 = new String[]{"20","21","22"};
        testScenarios.add(strings1);
        testScenarios.add(strings2);
        testScenarios.add(strings3);
        testScenarios.add(strings4);
        testScenarios.add(strings5);
        List<String> stringObjectOfTestCase = new LinkedList<>();
        stringObjectOfTestCase.add("ARTest-1");
        stringObjectOfTestCase.add("Event-1");
        List<List<String>> combinationOfScenarios = new LinkedList<>();
        ExcelReader excelReader = new ExcelReader();
        combinationOfScenarios = excelReader.getCombinationOfScenarios(testScenarios, stringObjectOfTestCase, combinationOfScenarios);
        System.out.println(combinationOfScenarios);
    }

    private void setAmountBasedOnBaseAmount(List<String> stringObjectOfTestCase) {
        if(stringObjectOfTestCase.size()==(rowLength-1)) {
            if(Integer.parseInt(stringObjectOfTestCase.get(5))>=100) {
                stringObjectOfTestCase.add("1");
            } else if(Integer.parseInt(stringObjectOfTestCase.get(5))<100) {
                stringObjectOfTestCase.add("0");
            } else {

            }
        }
    }
}
