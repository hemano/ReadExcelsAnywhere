package com.autmatika;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class LocalExcel extends ExcelLocationType {

    private String excelPath;
    private Workbook workbook;

    public LocalExcel(String excelPath) {
        this.excelPath = excelPath;
        try {
            FileInputStream excelFile = new FileInputStream(new File(this.excelPath));
            workbook = new XSSFWorkbook(excelFile);


        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.exit(0);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public LocalExcel(InputStream inputStream) {

        try {
            workbook = new XSSFWorkbook(inputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.exit(0);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    @Override
    public List<String> getExcelWorkSheetNames() throws IOException {
        List<String> sheetNamesList = new ArrayList<>();
        Iterator<Sheet> iterator = workbook.sheetIterator();
        while (iterator.hasNext()) {
            sheetNamesList.add(iterator.next().getSheetName());
        }
        return sheetNamesList;
    }

    @Override
    public List<List<Object>> getExcelData(String sheetName, String addressRangeOrUsedRange) throws IOException {

        List<List<Object>> tempRowsList = new ArrayList<>();
        List<Object> tempCellsList = new ArrayList<>();

        Sheet sheetObject = workbook.getSheet(sheetName);
        String reference = "";

        if (addressRangeOrUsedRange.equalsIgnoreCase("USEDRANGE")) {

            int lastRowNumber = sheetObject.getLastRowNum() + 1;
            int emptyRowsCount = sheetObject.getLastRowNum() + 1 - sheetObject.getPhysicalNumberOfRows();

            if (emptyRowsCount != 0) {
                lastRowNumber = lastRowNumber - emptyRowsCount - 1;
            }

            if (emptyRowsCount > 1) {
                throw new IOException("More than 1 empty rows are present. The last valid row number is: " + sheetObject.getLastRowNum());
            }

            Row firstRow = sheetObject.getRow(0);
            Cell lastCell = firstRow.getCell(firstRow.getLastCellNum() - 1);
            String lastCellAddress = lastCell.getAddress().formatAsString().replaceAll("[\\d+]", "");

            String range = "A1:" + lastCellAddress + lastRowNumber;
            reference = sheetName + "!" + range;
        } else {
            reference = sheetName + "!" + addressRangeOrUsedRange;
        }


        AreaReference aref = new AreaReference(reference, SpreadsheetVersion.EXCEL2007);
        CellReference[] crefs = aref.getAllReferencedCells();
        String cellValueString = "";


        for (int i = 0; i < crefs.length; i++) {

            Row r = sheetObject.getRow(crefs[i].getRow());
            Cell c = r.getCell(crefs[i].getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

            if (c != null) {
                switch (c.getCellTypeEnum()) {
                    case BLANK:
                        cellValueString = "";
                        break;
                    case STRING:
                        cellValueString = c.getStringCellValue();
                        break;
                    case NUMERIC:
                        cellValueString = (int) c.getNumericCellValue() + "";
                        break;
                    case ERROR:
//                        System.out.println("Error");
                        break;
                    case _NONE:
//                        System.out.println("None");
                        break;
                    case FORMULA:
//                        System.out.println("Formula");
                        break;
                    case BOOLEAN:
//                        System.out.println("Boolean");
                        break;
                    default:
//                        System.out.println("Default");

                }

                tempCellsList.add(cellValueString.trim());

            }
            if (i + 1 < crefs.length && crefs[i].getRow() != crefs[i + 1].getRow()) {
                tempRowsList.add(tempCellsList);
                tempCellsList = new ArrayList<>();
            }

            if (i == crefs.length - 1) {
                tempRowsList.add(tempCellsList);
            }
        }
        return tempRowsList;
    }


    @Override
    public void authorize() throws IOException {

    }

    @Override
    public Map<String, List<List<Object>>> getAllSheetsData(List<String> expectedSheetsList, String... workbookNames) throws IOException {
        return null;
    }


    public Map<String, List<List<Object>>> getAllSheetsData(List<String> exceptedSheets) throws IOException {

        List<String> sheets = getExcelWorkSheetNames();
        sheets.removeAll(exceptedSheets);
        Map<String, List<List<Object>>> sheetsDataMap = new HashMap<>();

        for(String sheet: sheets){
            sheetsDataMap.put(sheet, getExcelData(sheet, "UsedRange"));
        }

        return sheetsDataMap;
    }

}
