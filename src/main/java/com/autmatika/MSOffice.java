package com.autmatika;

import io.restassured.http.ContentType;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.stream.Collectors;

import static io.restassured.RestAssured.given;

public class MSOffice extends ExcelLocationType {

    private static final String TOKEN_BASE_URI = "https://login.microsoftonline.com/common/oauth2/v2.0";
    private static final String TOKEN_SCOPE = "files.readwrite offline_access sites.read.all";
    private static final String REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient";
    private static final String GRANT_TYPE = "refresh_token";


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

        return given()
                .header("Authorization", bearerToken)
                .baseUri(baseURI)
                .when()
                .get(path);
    }


    public List<List<Object>> getExcelData(String sheetName, String addressRangeOrUsedRange) {
        String path;
        //Decide to get either specified range or used range
        if (addressRangeOrUsedRange.equalsIgnoreCase("USEDRANGE")) {
            path = "/" + getResourceId() + "/workbook/worksheets/" + sheetName + "/UsedRange";
        } else {
            path = "/" + getResourceId() + "/workbook/worksheets/" + sheetName + "/range(address='" + addressRangeOrUsedRange + "')";
        }

        String response = executeGetRequest(path).asString();

        JsonPath jsonPath = new JsonPath(response);

        return jsonPath.getList("formulas");
    }

}

