package com.autmatika;

import com.jayway.jsonpath.Filter;
import io.restassured.http.ContentType;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static com.jayway.jsonpath.Criteria.where;
import static com.jayway.jsonpath.Filter.filter;
import static io.restassured.RestAssured.given;

public class MSOffice extends ExcelLocationType {

    private static final String TOKEN_BASE_URI = "https://login.microsoftonline.com/common/oauth2/v2.0";
    private static final String TOKEN_SCOPE = "files.readwrite offline_access sites.read.all";
    private static final String REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient";
    private static final String GRANT_TYPE = "refresh_token";

    private static Map<String, List<List<Object>>> sharedSheetsAndDataMap;

    private String baseURI;
    private ExcelLocation excelLocation;
    private LocalExcel localExcel;


    //Read this values from properties
    private String applicationId;
    private String refreshToken;
    private String accessToken;
    private String resourceId;

    public MSOffice(ExcelLocation excelLocation, String applicationId, String resourceId, String refreshToken) {
        this.excelLocation = excelLocation;
        this.resourceId = resourceId;
        this.applicationId = applicationId;
        this.refreshToken = refreshToken;

        switch (excelLocation) {
            case ONE_DRIVE:
                baseURI = "https://graph.microsoft.com/v1.0/me/drive/items";
                break;
            case SHAREPOINT:
                baseURI = "https://graph.microsoft.com/v1.0/sites/root/drive/items";
                break;
        }
        authorize();
    }

    public MSOffice(ExcelLocation excelLocation, String pathOfLocalExcelFile) {
        this.excelLocation = excelLocation;
        this.localExcel = new LocalExcel(pathOfLocalExcelFile);
    }


    private String getAccessToken() {
        return accessToken;
    }

    private String getApplicationId() {
        return applicationId;
    }

    private String getRefreshToken() {
        return refreshToken;
    }

    public String getResourceId() {
        return resourceId;
    }

    public Map<String, List<List<Object>>> getSharedSheetsAndDataMap() {
        return sharedSheetsAndDataMap;
    }

    public void authorize() {

        String responseAsString;
        Response response;
        try {
            response = given()
                    .contentType(ContentType.URLENC)
                    .baseUri(TOKEN_BASE_URI)
                    .param("client_id", getApplicationId())
                    .param("scope", TOKEN_SCOPE)
                    .param("refresh_token", getRefreshToken())
                    .param("redirect_uri", REDIRECT_URI)
                    .param("grant_type", GRANT_TYPE)
                    .when()
                    .post("/token");

            responseAsString = response.asString();
            JsonPath jsonPath = new JsonPath(responseAsString);

            if (response.getStatusCode() != 200) {
                throw new IllegalArgumentException("Error: " + jsonPath.get("error") + " Error description: " + jsonPath.get("error_description"));
            }

            this.refreshToken = jsonPath.getString("refresh_token");
            this.accessToken = jsonPath.getString("access_token");

        } catch (IllegalArgumentException e) {
            e.printStackTrace();
        }

    }


    public List<String> getExcelWorkSheetNames() throws IOException {

        if (this.excelLocation.equals(ExcelLocation.LOCAL)) {
            return localExcel.getExcelWorkSheetNames();
        } else {
            String path = getResourceId() + "/workbook/worksheets";

            String response = executeGetRequest(path).asString();

            JsonPath jsonPath = new JsonPath(response);

            return jsonPath.getList("value").stream().map(e -> ((HashMap<String, String>) e).get("name")).collect(Collectors.toList());
        }

    }


    /**
     * @param path This is a path of the graph api followed by DRIVE_BASE_URI
     * @return Returns the response object of RestAssured Api
     */
    private Response executeGetRequest(String path) {
        String bearerToken = "Bearer " + getAccessToken();

        Response response = given()
                .header("Authorization", bearerToken)
                .baseUri(baseURI)
                .when()
                .get(path);


        String responseString = response.asString();
        JsonPath jsonPath = new JsonPath(responseString);

        if (response.getStatusCode() != 200) {
            throw new IllegalArgumentException("Error: " + jsonPath.get("error") + " Error description: " + jsonPath.get("error_description"));
        }

        return response;

    }


    public List<List<Object>> getExcelData(String sheetName, String addressRangeOrUsedRange) throws IOException {

        System.out.println("Reading Data from sheet: " + sheetName);
        if (ExcelLocation.LOCAL.equals(excelLocation)) {
            return localExcel.getExcelData(sheetName, addressRangeOrUsedRange);
        } else {
            Map<String, List<List<Object>>> map = getSharedSheetsAndDataMap();
            Map<String, List<List<Object>>> mapOfSheetsAndData = new HashMap<>();

            for (Map.Entry entry : map.entrySet()) {
                mapOfSheetsAndData.put(entry.getKey().toString(), (List<List<Object>>) entry.getValue());
            }
            return mapOfSheetsAndData.get(sheetName);
        }
    }

    public Map<String, List<List<Object>>> getAllSheetsData(List<String> expectedSheetsList, String... workbookNames) throws IOException {

        // TODO: 02/02/18 Handle the condition of no workbook names
        Map<String, String> workbooksAndIDMap = getMapOfWorkbookAndIDs(workbookNames);
        Map<String, List<List<Object>>> bigMap = new HashMap<>();


        // TODO: 05/02/18
        for (String workBook : workbookNames) {

            resourceId = workbooksAndIDMap.get(workBook);

            Response response = new RestUtil().executeGetRequest(baseURI, getAccessToken(), resourceId + "/content");

            LocalExcel localExcel = new LocalExcel(response.asInputStream());

            Map<String, List<List<Object>>> map = localExcel.getAllSheetsData(expectedSheetsList);

            bigMap.putAll(map);
        }
        sharedSheetsAndDataMap = bigMap;
        return bigMap;


//        JSONArray jsonArray = new JSONArray();
//        JSONObject jsonObject = new JSONObject();
//
//        int idCount = 1;
//        for (Map.Entry entry : workbooksAndIDMap.entrySet()) {
//            JSONObject tempJsonObject = new JSONObject();
//
//            String tempUrl = "sites/root/drive/items/" + entry.getValue() + "/workbook/worksheets";
//
//            tempJsonObject.put("url", tempUrl);
//            tempJsonObject.put("method", "GET");
//            tempJsonObject.put("id", idCount++);
//
//            jsonArray.add(tempJsonObject);
//        }
//
//        jsonObject.put("requests", jsonArray);
//        String batchRequestBooksAndSheetsString = jsonObject.toJSONString().replace("\\", "");
//
////        Path filePath = new ReadWriteUtil().writeFilesInTargetFolder("jsons",
////                "sheetsRequestBatch.json",
////                batchRequestBooksAndSheetsString.getBytes());
//
//        org.asynchttpclient.Response batchResponseBooksAndSheets = new RestUtil().executePostRequest("https://graph.microsoft.com/beta/$batch", getAccessToken(), batchRequestBooksAndSheetsString);
//
////        Path sheetsResponseBatchPath = new ReadWriteUtil().writeFilesInTargetFolder("jsons", "sheetsResponseBatch.json", batchResponseBooksAndSheets.asByteArray());
//
//        Map<String, List<String>> mapOfWorkBookAndSheets = new HashMap<>();
//
////        String sheetsResponseFromFile = new ReadWriteUtil().readFileFromTargerFolder(sheetsResponseBatchPath);
//
//        int index = 1;
//        for (Map.Entry s : workbooksAndIDMap.entrySet()) {
//            Filter workbooksFilter = filter(where("id").is(String.valueOf(index++)));
//            List<String> list = com.jayway.jsonpath.JsonPath.parse(batchResponseBooksAndSheets.getResponseBody()).read( "responses[?].body.value[*].name", workbooksFilter);
//            mapOfWorkBookAndSheets.put(s.getKey().toString(), list);
//        }
//
//        //Create batch request to get all the data from sheets
//
//        jsonArray = new JSONArray();
//        jsonObject = new JSONObject();
//
//        idCount = 1;
//        List<String> listOfSheet = new ArrayList<>();
//
//        for (Map.Entry entry : mapOfWorkBookAndSheets.entrySet()) {
//
//            for (String sheetName : (List<String>) entry.getValue()) {
//
//                if(!sheetName.equalsIgnoreCase("test")
//                        && !sheetName.equalsIgnoreCase("ActionsDescription")
//                        && !sheetName.equalsIgnoreCase("Cleanup")
//                        && !sheetName.equalsIgnoreCase("IACC")
//                        && !sheetName.equalsIgnoreCase("Banner")){
//                    JSONObject tempJsonObject = new JSONObject();
//
//                    String tempUrl = "sites/root/drive/items/" + workbooksAndIDMap.get(entry.getKey()) + "/workbook/worksheets/" + sheetName + "/usedrange(valuesOnly=true)";
//
//                    tempJsonObject.put("url", tempUrl);
//                    tempJsonObject.put("method", "GET");
//                    tempJsonObject.put("id", idCount++);
//
//                    jsonArray.add(tempJsonObject);
//                    listOfSheet.add(sheetName);
//                }
//            }
//
//        }
//
//        jsonObject.put("requests", jsonArray);
//        String batchRequestSheetsAndData = jsonObject.toJSONString().replace("\\", "");
//
//
////        Path batchRequestSheetsAndDataPath = new ReadWriteUtil().writeFilesInTargetFolder("jsons",
////                "batchRequestSheetsAndData.json",
////                batchRequestSheetsAndData.getBytes());
//
//        org.asynchttpclient.Response batchResponseSheetsAndData = new RestUtil().executePostRequest("https://graph.microsoft.com/beta/$batch", getAccessToken(), batchRequestSheetsAndData);
////        Path batchResponseSheetsAndDataPath = new ReadWriteUtil().writeFilesInTargetFolder("jsons", "batchResponseSheetsAndData.json", batchResponseSheetsAndData.asByteArray());
//
////        String batchResponseSheetsAndDataFromFile = new ReadWriteUtil().readFileFromTargerFolder(batchResponseSheetsAndDataPath);
//
//
//        Map<String, List<Object>> sheetsAndDataMap = new HashMap<>();
//
//        for (int i = 1; i <= listOfSheet.size(); i++) {
//            Filter workbooksFilter = filter(where("id").is(String.valueOf(i)));
//            List<Object> list = com.jayway.jsonpath.JsonPath.parse(batchResponseSheetsAndData.getResponseBody()).read( "responses[?].body.formulas", workbooksFilter);
//            List<String> list1 = com.jayway.jsonpath.JsonPath.parse(batchResponseSheetsAndData.getResponseBody()).read( "responses[?].body.address", workbooksFilter);
//            String sheetName = list1.get(0).toString().substring(0, list1.get(0).toString().indexOf("!")).replaceAll("'","");
//            sheetsAndDataMap.put(sheetName, (List<Object>)list.get(0));
//        }

//        sharedSheetsAndDataMap = sheetsAndDataMap;

//        return null;
    }

    public Map<String, String> getMapOfWorkbookAndIDs(String... workbookNames) throws IOException {

        RestUtil restUtil = new RestUtil();
        Map<String, String> mapOfSheetAndIds = new HashMap<>();

        String endPointPath = "https://graph.microsoft.com/v1.0/sites/root/drive/root/search(q='.xlsx')?select=name,id,webUrl";
        Response response = restUtil.executeGetRequest("", getAccessToken(), endPointPath);

        for (String bookName : workbookNames) {

            Filter workbooksFilter = filter(where("name").is(bookName));
            List<String> workBookID = com.jayway.jsonpath.JsonPath.parse(response.asInputStream()).read(
                    "$.value[?].id",
                    workbooksFilter);
            mapOfSheetAndIds.put(bookName, workBookID.get(0));
        }

        return mapOfSheetAndIds;
    }
}

