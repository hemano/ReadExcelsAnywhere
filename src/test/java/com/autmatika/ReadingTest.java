package com.autmatika;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.time.Duration;
import java.time.LocalTime;
import java.util.*;
import java.util.stream.Collectors;


public class ReadingTest {

    private String googleApiKey;
    private String graphRefreshToken;

    @BeforeSuite
    public void before() throws IOException {
        String configFilePath = getClass().getClassLoader().getResource("config.properties").getPath();
        Properties properties = new Properties();

        properties.load(new InputStreamReader(new FileInputStream(new File(configFilePath))));

        googleApiKey = properties.getProperty("google_api_key");
        graphRefreshToken = properties.getProperty("microsoft_graph_refresh_token");
    }

    @Test(enabled = false)
    public void testGoogleSheetViaOAuth() throws Exception {

        String googleSheetResourceId = googleApiKey;

        GoogleDriveOAuth googleDriveOAuth = new GoogleDriveOAuth(googleSheetResourceId);
        ReadExcel<GoogleDriveOAuth> readExcel = new ReadExcel<>();

        List<String> sheets = readExcel.getListOfSheets(googleDriveOAuth);

        System.out.println(sheets);

    }

    @Test(enabled = false)
    public void testGoogleSheetRead() throws Exception {

        String resourceId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
        String key = googleApiKey;

        GoogleDriveAPI googleDriveAPI = new GoogleDriveAPI(key, resourceId);
        ReadExcel<GoogleDriveAPI> readExcel = new ReadExcel<>();

        List<String> sheets = readExcel.getListOfSheets(googleDriveAPI);

        System.out.println(sheets);
    }

    @Test(enabled = false)
    public void testGoogleSheetReadData() throws Exception {

        //https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit#gid=0
        String resourceId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
        String key = googleApiKey;
        String sheetName = "Class Data";

        GoogleDriveAPI googleDriveAPI = new GoogleDriveAPI(key, resourceId);
        ReadExcel<GoogleDriveAPI> readExcel = new ReadExcel<>();

        List<List<Object>> data = readExcel.getExcelData(googleDriveAPI, sheetName, "UsedRange");
        System.out.println(data);

        String[][] dataArray = readExcel.getExcelDataInStringArray(googleDriveAPI, "UsedRange");
        System.out.println(dataArray);

    }


    @Test(enabled = true)
    public void testSharePointExcelAnotherWorkbook() throws Exception {

        String applicationId = "bb3435a1-869c-494b-9cb0-793f145dd316";
        String refreshToken = graphRefreshToken;
        String resourceId = "01XI34BXAZOXQNPB56TNHKPB3PVYHDEEBB";

        MSOffice msOffice = new MSOffice(ExcelLocation.SHAREPOINT, applicationId, resourceId, refreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

//        List<String> sheets = readExcel.getListOfSheets(msOffice);
//        System.out.println(sheets);


//        List<List<Object>> data = readExcel.getExcelData(msOffice, "ActionsDescription", "UsedRange");
//        System.out.println(data);


        Map<String, List<ArrayList<String>>> map = readExcel.getMapOfSheetsAndData(msOffice, "UsedRange", Arrays.asList(""));
        System.out.println(map);

    }

    @Test(enabled = true)
    public void adhocTest() throws Exception {

        Map<String, List<String>> sheetAndTestMap = new LinkedHashMap<>();

        MSOffice msOffice = new MSOffice(
                ExcelLocation.SHAREPOINT,
                "bb3435a1-869c-494b-9cb0-793f145dd316",
                "01XI34BXD75CFWGRX5QNCKLQBYZF65VCLE",
                graphRefreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

        Map<String, List<ArrayList<String>>> sheetsDataMap = readExcel.getMapOfSheetsAndData(msOffice,
                "UsedRange",
                java.util.Arrays.asList("Test"));

        Map<String, String> flagAndTest = new HashMap<>();

        for (Map.Entry entry : sheetsDataMap.entrySet()) {

            List<ArrayList<String>> rowsInASheet = (ArrayList) entry.getValue();

            flagAndTest = new HashMap<>();
            for (ArrayList<String> row : rowsInASheet) {

                if (!row.get(1).trim().equals(""))
                    flagAndTest.put(row.get(1), row.get(0));
            }

            //removing the data of first row of every sheet
            flagAndTest.remove("Test");

            ArrayList<Map.Entry> list = (ArrayList) flagAndTest.entrySet().stream().filter(r -> r.getValue().equals("Y")).collect(Collectors.toList());

            List<String> listOfTests = list.stream().map(Map.Entry::getKey).collect(Collectors.toList()).stream().map(Object::toString).collect(Collectors.toList());

            sheetAndTestMap.put(entry.getKey().toString(), listOfTests);
        }

        System.out.println();
    }

    @Test(enabled = true)
    public void adhocTest1() throws Exception {

        MSOffice msOffice = new MSOffice(
                ExcelLocation.SHAREPOINT,
                "bb3435a1-869c-494b-9cb0-793f145dd316",
                "01XI34BXAP6ARY7JPTJVCINGPJZ5KA7OE7",
                graphRefreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

        List<ArrayList<String>> temp = readExcel.getExcelDataInStringFormat(msOffice,
                "Properties",
                "UsedRange");

        System.out.println();
    }

    @Test(enabled = true)
    public void adhocTest2() throws Exception {

        Map<String, List<String>> sheetAndTestMap = new LinkedHashMap<>();

        MSOffice msOffice = new MSOffice(
                ExcelLocation.SHAREPOINT,
                "bb3435a1-869c-494b-9cb0-793f145dd316",
                "01XI34BXD75CFWGRX5QNCKLQBYZF65VCLE",
                graphRefreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

        String[][] temp = readExcel.getExcelDataInStringArray(msOffice, "UsedRange", Arrays.asList(""));

        Map<String, List<ArrayList<String>>> sheetsDataMap = readExcel.getMapOfSheetsAndData(msOffice,
                "UsedRange",
                java.util.Arrays.asList("Test"));

    }

    @Test(enabled = true)
    public void testBatchRequest() throws Exception {

        MSOffice msOffice = new MSOffice(
                ExcelLocation.SHAREPOINT,
                "bb3435a1-869c-494b-9cb0-793f145dd316",
                "01XI34BXD75CFWGRX5QNCKLQBYZF65VCLE",
                graphRefreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

//        String[][] temp =readExcel.getExcelDataInStringArray(msOffice,"UsedRange",Arrays.asList(""));
//        System.out.println(true);
//
//        Map<String, List<ArrayList<String>>> sheetsDataMap = readExcel.getMapOfSheetsAndData(msOffice,
//                "UsedRange",
//                java.util.Arrays.asList("Test"));

    }


    @Test
    public void createBatchRequestToFetchAllData() throws IOException {

        LocalTime time1 = LocalTime.now();
        JSONObject batchRequestForSheets = new JSONObject();
        JSONArray jsonArray = new JSONArray();

        Map<String, String> workBookNamesAndIDsMap = new HashMap<>();

        MSOffice sharePointAccess = new MSOffice(
                ExcelLocation.SHAREPOINT,
                "bb3435a1-869c-494b-9cb0-793f145dd316",
                "01XI34BXD75CFWGRX5QNCKLQBYZF65VCLE",
                graphRefreshToken);

//        Map<String, List<Object>> ma = sharePointAccess.getAllSheetsData("Test.xlsx");

        LocalTime time2 = LocalTime.now();
        System.out.println(Duration.between(time1, time2).toNanos());
        System.out.println();

    }


    @Test(enabled = false)
    public void testOneDriveExcel() throws Exception {

        String applicationId = "d1c318de-dcb6-4e35-a1ad-15907c7b8744";
        String refreshToken = graphRefreshToken;
        String resourceId = "d1c318de-dcb6-4e35-a1ad-15907c7b8744";

        MSOffice office = new MSOffice(ExcelLocation.ONE_DRIVE, applicationId, resourceId, refreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

        List<String> sheets = readExcel.getListOfSheets(office);
        System.out.println(sheets);
    }

    @Test
    public void testLocalExcel() throws Exception {
        String localExcelPath = getClass().getClassLoader().getResource("SmokeTests.xlsx").getPath();
        MSOffice office = new MSOffice(ExcelLocation.LOCAL, localExcelPath);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

        List<String> sheets = readExcel.getListOfSheets(office);
        System.out.println(sheets);
    }

    @Test
    public void testLocalExcelIfItBringCellsOnRange() throws Exception {
        String localExcelPath = getClass().getClassLoader().getResource("SmokeTests.xlsx").getPath();
        MSOffice office = new MSOffice(ExcelLocation.LOCAL, localExcelPath);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

        List<List<Object>> list = readExcel.getExcelData(office, "SmokeTest", "UsedRange");
        System.out.println(list);

//        readExcel.getExcelDataInStringArray(office,"UsedRange");

        readExcel.getExcelDataInStringArray(office, "UsedRange", Arrays.asList("DoNotConsider"));
    }

    @Test
    public void testAnotherLocalExcelIfItBringCellsOnRange() throws Exception {
        String localExcelPath = getClass().getClassLoader().getResource("BusinessProcess.xlsx").getPath();
        MSOffice office = new MSOffice(ExcelLocation.LOCAL, localExcelPath);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

//        List<List<Object>> list = readExcel.getExcelData(office, "BusinessProcess", "UsedRange");
//        System.out.println(list);

        String[][] data = readExcel.getExcelDataInStringArray(office, "UsedRange", Arrays.asList("ActionsDescription"));

//        String[][] data = readExcel.getExcelDataInStringArray(office, "UsedRange", Arrays.asList("DoNotConsider"));
        System.out.println(data);
    }


    @Test
    public void readLocatorUIFromSharePoint() throws IOException {
        MSOffice msOffice = new MSOffice(
                ExcelLocation.SHAREPOINT,
                "bb3435a1-869c-494b-9cb0-793f145dd316",
                "01XI34BXAP6ARY7JPTJVCINGPJZ5KA7OE7",
                graphRefreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>(msOffice, new String[]{"LocatorPropertiesUI.xlsx"});

        List<ArrayList<String>> temp = readExcel.getExcelDataInStringFormat(msOffice,
                "Properties",
                "UsedRange");

        System.out.println(temp);

        List<ArrayList<String>> temp1 = readExcel.getExcelDataInStringFormat(msOffice,
                "Base",
                "UsedRange");

        System.out.println(temp1);

        List<ArrayList<String>> temp2 = readExcel.getExcelDataInStringFormat(msOffice,
                "Properties",
                "UsedRange");

        List<ArrayList<String>> temp4 = readExcel.getExcelDataInStringFormat(msOffice,
                "Base",
                "UsedRange");

//        List<ArrayList<String>> temp5 = temp2.addAll(temp4);
        temp2.addAll(temp4);

        System.out.println(temp2);

    }


    @Test
    public void readTestsFromSharePoint() throws IOException {
        MSOffice msOffice = new MSOffice(
                ExcelLocation.SHAREPOINT,
                "bb3435a1-869c-494b-9cb0-793f145dd316",
                "",
                graphRefreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>(msOffice, new String[]{"Test.xlsx"});

        String[][] temp = readExcel.getExcelDataInStringArray(msOffice, "UsedRange", Arrays.asList("Test"));

        System.out.println(Arrays.toString(temp));
    }

    @Test
    public void readLocatorPropertiesUIFromSharePoint() throws IOException {
        MSOffice msOffice = new MSOffice(
                ExcelLocation.SHAREPOINT,
                "bb3435a1-869c-494b-9cb0-793f145dd316",
                "",
                graphRefreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>(msOffice, new String[]{"LocatorPropertiesUI.xlsx"});

        List<ArrayList<String>> temp2 = readExcel.getExcelDataInStringFormat(msOffice,
                "Properties",
                "UsedRange");

        System.out.println(temp2);
    }


    @Test
    public void readBusinessProcessFromSharePoint() throws IOException {

        MSOffice msOffice = new MSOffice(
                ExcelLocation.SHAREPOINT,
                "bb3435a1-869c-494b-9cb0-793f145dd316",
                "01XI34BXAP6ARY7JPTJVCINGPJZ5KA7OE7",
                graphRefreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>(msOffice, new String[]{"BusinessProcess.xlsx"}, Arrays.asList("ActionsDescription"));

        String[][] arr = readExcel.getExcelDataInStringArray(msOffice,
                "UsedRange",
                java.util.Arrays.asList("ActionsDescription"));

        System.out.println(Arrays.toString(arr));

    }

    @Test
    public void getMapOfSheetsAndTestsFromSharePoint() throws IOException {

        MSOffice msOffice = new MSOffice(
                ExcelLocation.SHAREPOINT,
                "bb3435a1-869c-494b-9cb0-793f145dd316",
                "",
                graphRefreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>(msOffice, new String[]{"Test.xlsx"});

        Map<String, List<ArrayList<String>>> sheetsDataMap = readExcel.getMapOfSheetsAndData(msOffice, "UsedRange", Arrays.asList(""));

        System.out.println(sheetsDataMap);

    }
}
