package com.autmatika;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.Spreadsheet;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class GoogleDriveOAuth extends ExcelLocationType {

    private static GoogleDriveOAuth instance;

    private final String APPLICATION_NAME = "Google Sheets API Quickstart";
    private final File DATA_STORE_DIR = new File(
            System.getProperty("user.dir"), "/src/main/resources/googleCredentials");
    private FileDataStoreFactory DATA_STORE_FACTORY;
    private final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private HttpTransport HTTP_TRANSPORT;
    private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS_READONLY);

    private Credential credential;
    private String resourceId;

    public String getResourceId() {
        return resourceId;
    }


    public GoogleDriveOAuth(String resourceId) throws IOException {
        this.resourceId = resourceId;

        try {
            if (GoogleDriveOAuth.class.getResourceAsStream("/client_secret.json") == null ||
                    !new File(DATA_STORE_DIR+"/credentials").exists()) {
                    throw new FileNotFoundException("client_secret.json or credentials doesn't exists");
            }

            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
        } catch (
                Throwable e)

        {
            e.printStackTrace();
            System.exit(1);
        }

        //Authorize and create credentials
        authorize();

    }


    /**
     * Creates and authorized Credential object.
     *
     * @return and authorized Credential object.
     * @throws IOException
     */
    public void authorize() throws IOException {

        //Load client secrets.
        InputStream in = GoogleDriveOAuth.class.getResourceAsStream("/client_secret.json");
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow =
                new GoogleAuthorizationCodeFlow.Builder(
                        HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                        .setDataStoreFactory(DATA_STORE_FACTORY)
                        .setAccessType("offline")
                        .build();

        Credential credential = new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver()).authorize("user");

        System.out.println(
                "Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());

        this.credential = credential;
    }

    @Override
    public Map<String, List<List<Object>>> getAllSheetsData(List<String> expectedSheetsList, String... workbookNames) throws IOException {
        return null;
    }


    public Sheets getSheetService() throws IOException {

        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, this.credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }


    public List<String> getExcelWorkSheetNames() throws IOException {
        List<String> ranges = new ArrayList<>();
        boolean includeGridData = false;
        Spreadsheet response = getSheetService().spreadsheets().get(getResourceId()).setRanges(ranges).setIncludeGridData(includeGridData).execute();
        return response.getSheets().stream().map(e -> e.getProperties().getTitle()).collect(Collectors.toList());
    }


    public List<List<Object>> getExcelData(String sheetName, String addressRangeOrUsedRange) throws IOException {

        String range;

        //Decide to get either specified range or used range
        if (addressRangeOrUsedRange.equalsIgnoreCase("USEDRANGE")) {
            range = sheetName;
        } else {
            range = sheetName + "!" + addressRangeOrUsedRange;
        }


        return getSheetService()
                .spreadsheets()
                .values()
                .get(getResourceId(), range)
                .execute()
                .getValues();

    }

}
