package com.autmatika;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class LocalExcel extends ExcelLocationType {

    private String excelPath;
    private Workbook workbook;

    public LocalExcel(String excelPath){
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
    @Override
    public List<String> getExcelWorkSheetNames() throws IOException {
        List<String> sheetNamesList = new ArrayList<>();
        Iterator<Sheet> iterator = workbook.sheetIterator();
        while(iterator.hasNext()){
            sheetNamesList.add(iterator.next().getSheetName());
        }
        return sheetNamesList;
    }

    @Override
    public List<List<Object>> getExcelData(String sheetName, String addressRangeOrUsedRange) throws IOException {
        return null;
    }

    @Override
    public void authorize() throws IOException {

    }
}
