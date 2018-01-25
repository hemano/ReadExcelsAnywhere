package com.autmatika;

import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

import java.io.*;
import java.util.List;
import java.util.Properties;


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

    @Test(enabled = false)
    public void testSharePointExcel() throws Exception {
        String applicationId = "bb3435a1-869c-494b-9cb0-793f145dd316";
        String refreshToken = graphRefreshToken;
        String resourceId = "01XI34BXCOA2S5NQIA3JEIFJJTG5AM3KY5";

        MSOffice msOffice = new MSOffice(ExcelLocation.SHAREPOINT, applicationId, resourceId, refreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

        List<String> sheets = readExcel.getListOfSheets(msOffice);
        System.out.println(sheets);

        List<List<Object>> data = readExcel.getExcelData(msOffice, "Salesforce", "UsedRange");
        System.out.println(data);
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

        List<List<Object>> list = readExcel.getExcelData(office, "Sheet2", "UsedRange");
        System.out.println(list);

        readExcel.getExcelDataInStringArray(office,"UsedRange");
    }


}
