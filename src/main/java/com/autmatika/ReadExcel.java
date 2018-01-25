package com.autmatika;

import org.apache.commons.lang3.ArrayUtils;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ReadExcel<T extends ExcelLocationType> {

    public List<String> getListOfSheets(T service) throws IOException {
        return service.getExcelWorkSheetNames();
    }

    public List<List<Object>> getExcelData(T service, String sheetName, String addressRangeOrUsedRange ) throws IOException {
        return service.getExcelData( sheetName, addressRangeOrUsedRange);
    }

    public List<ArrayList<String>> getExcelDataInStringFormat(T service,  String sheetName, String addressRangeOrUsedRange) throws IOException {

        List<List<Object>> tempList = getExcelData(service, sheetName, addressRangeOrUsedRange);

        List<ArrayList<String>> tempList1 = new ArrayList<>();

        for (List<Object> objects : tempList) {
            tempList1.add((ArrayList<String>) (ArrayList<?>) (objects));
        }
        return tempList1;
    }

    public String[][] getExcelDataInStringArray(T service, String sheetName, String addressRangeOrUsedRange) throws IOException {

        List<ArrayList<String>> values = getExcelDataInStringFormat(service, sheetName, addressRangeOrUsedRange);

        return getStrings(values);
    }


    public String[][] getExcelDataInStringArray(T service, String addressRangeOrUsedRange, List<String> exceptedSheetsList) throws IOException {

        List<String[][]> listOfTables = new ArrayList<>();

        List<String> sheets = getListOfSheets(service);
        sheets.removeAll(exceptedSheetsList);

        boolean firstRow = true;
        for (String sheet : sheets) {

            //Removing the first row from the subsequent tables/ sheets to keep header row only once
            if (firstRow == true) {
                firstRow = false;
                listOfTables.add(getStrings(getExcelDataInStringFormat(service, sheet, addressRangeOrUsedRange)));
            } else {
                String[][] src = getStrings(getExcelDataInStringFormat(service, sheet, addressRangeOrUsedRange));
                String[][] dest = new String[src.length - 1][src[0].length];

                for (int i = 1; i < src.length; i++) {
                    System.arraycopy(src[i], 0, dest[i - 1], 0, src[0].length - 1);
                }
                listOfTables.add(dest);
            }


        }

        String[][] temp = null;
        for (String[][] table : listOfTables) {
            temp = (String[][]) ArrayUtils.addAll(temp, table);
        }
        return temp;
    }

    public String[][] getExcelDataInStringArray(T service, String addressRangeOrUsedRange) throws IOException {

        List<String[][]> listOfTables = new ArrayList<>();

        List<String> sheets = getListOfSheets(service);

        boolean firstRow = true;
        for (String sheet : sheets) {

            //Removing the first row from the subsequent tables/ sheets to keep header row only once
            if (firstRow == true) {
                firstRow = false;
                listOfTables.add(getStrings(getExcelDataInStringFormat(service, sheet, addressRangeOrUsedRange)));
            } else {
                String[][] src = getStrings(getExcelDataInStringFormat(service, sheet, addressRangeOrUsedRange));
                String[][] dest = new String[src.length - 1][src[0].length];

                for (int i = 1; i < src.length; i++) {
                    System.arraycopy(src[i], 0, dest[i - 1], 0, src[0].length - 1);
                }
                listOfTables.add(dest);
            }


        }

        String[][] temp = null;
        for (String[][] table : listOfTables) {
            temp = (String[][]) ArrayUtils.addAll(temp, table);
        }
        return temp;
    }


    private String[][] getStrings(List<ArrayList<String>> values) throws IOException {
        String[][] temp;

        if (values == null || values.size() == 0) {
            throw new IOException("No Data Found");
        } else {
            temp = new String[values.size()][values.get(0).size()];
            int rowIndex = 0;
            for (List row : values) {
                for (int columnIndex = 0; columnIndex < values.get(0).size(); columnIndex++) {
                    try {
                        temp[rowIndex][columnIndex] = (String) row.get(columnIndex);
                    } catch (Exception e) {
                        temp[rowIndex][columnIndex] = "";
                    }

                }
                rowIndex++;
            }
        }
        return temp;
    }
}
