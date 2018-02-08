package com.autmatika;

import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import static io.restassured.RestAssured.given;

public class GoogleDriveAPI extends ExcelLocationType {

    private static final String BASE_URI = "https://sheets.googleapis.com/v4/";
    private String key;
    private String resourceId;

    public GoogleDriveAPI(String key, String resourceId) {
        this.key = key;
        this.resourceId = resourceId;
    }

    public String getKey() {
        return key;
    }

    public String getResourceId() {
        return resourceId;
    }

    @Override
    public List<String> getExcelWorkSheetNames() throws IOException {

        String url = "/spreadsheets/" + getResourceId();

        String responseAsString = null;
        try {

            Response response = given()
                    .baseUri(BASE_URI)
                    .param("key", getKey())
                    .when()
                    .get(url);

            responseAsString = response.asString();
            JsonPath jsonPath = new JsonPath(responseAsString);

            if (response.statusCode() != 200) {
                throw new IllegalArgumentException("Code: " + jsonPath.get("error.code") + " Message: " + jsonPath.get("error.message"));
            }
            return jsonPath.get("sheets.properties.title");
        } catch (IllegalArgumentException e) {
            e.printStackTrace();
            return null;
        }

    }

    @Override
    public List<List<Object>> getExcelData(String sheetName, String addressRangeOrUsedRange) throws IOException {
        String range;

        if (addressRangeOrUsedRange.equalsIgnoreCase("USEDRANGE")) {
            range = sheetName;
        } else {
            range = sheetName + "!" + addressRangeOrUsedRange;
        }

        String url = "/spreadsheets/" + getResourceId() + "/values/" + range;

        String response = given()
                .baseUri(BASE_URI)
                .param("key", getKey())
                .when()
                .get(url).asString();

        JsonPath jsonPath = new JsonPath(response);

        return jsonPath.getList("values");

    }

    @Override
    public void authorize() throws IOException {

    }

    @Override
    public Map<String, List<List<Object>>> getAllSheetsData(List<String> expectedSheetsList, String... workbookNames) throws IOException {
        return null;
    }

}
