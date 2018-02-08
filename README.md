# ReadExcelsAnywhere
<h2>To obtain an api key for Google sheets</h2>

<ol>
<li>Visit https://console.developers.google.com/apis</li>
<li>Create a project and go to Dashboard</li>
<li>Enable Google Sheets API</li>
<li>Go to Credentials</li>
<li>Click Create Credentials and select API key</li>
<li>Copy the API key showed in the dialog box.</li>
</ol>
<hr>

<h2>To get access to Excel sheet stored on Sharepoint or OneDrive</h2>
<ol>
<li>Visit https://developer.microsoft.com/en-us/graph</li>
<li>Login with a valid account</li>
<li>Follow the instructions to register your app and make a note of the bearer token </li>
<li>https://developer.microsoft.com/en-us/graph/docs/concepts/auth_v2_service</li>
<li>Follow below instruction to generate a refresh token</li>
<li>https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols-oauth-code</li>
<li>Refresh will enable user to access the resources without logging in again/ bearer token expires</li>
</ol>

<hr>
<h2> Sample Tests </h2>

    @Test(enabled = false)
    public void testGoogleSheetViaOAuth() throws IOException {

        String googleSheetResourceId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";

        GoogleDriveOAuth googleDriveOAuth = new GoogleDriveOAuth(googleSheetResourceId);
        ReadExcel<GoogleDriveOAuth> readExcel = new ReadExcel<>();

        List<String> sheets = readExcel.getListOfSheets(googleDriveOAuth);
        System.out.println(sheets);
    }

    @Test
    public void testGoogleSheetRead() throws IOException {

        String resourceId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
        String key = "<api_key>";

        GoogleDriveAPI googleDriveAPI = new GoogleDriveAPI(key,resourceId);
        ReadExcel<GoogleDriveAPI> readExcel = new ReadExcel<>();

        List<String> sheets = readExcel.getListOfSheets(googleDriveAPI);

        System.out.println(sheets);
    }

    @Test
    public void testGoogleSheetReadData() throws IOException {

        //https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit#gid=0
        String resourceId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
        String key = "<api_key>";
        String sheetName = "Class Data";

        GoogleDriveAPI googleDriveAPI = new GoogleDriveAPI(key, resourceId);
        ReadExcel<GoogleDriveAPI> readExcel = new ReadExcel<>();

        List<List<Object>> data = readExcel.getExcelData(googleDriveAPI, sheetName, "UsedRange");
        System.out.println(data);

        String[][] dataArray = readExcel.getExcelDataInStringArray(googleDriveAPI,"UsedRange");
        System.out.println(dataArray);

    }

    @Test
    public void testSharePointExcel() throws IOException {
        String applicationId = "bb3435a1-869c-494b-9cb0-793f145dd316";
        String refreshToken = "<refresh_token>";
        String resourceId = "01XI34BXCOA2S5NQIA3JEIFJJTG5AM3KY5";

        MSOffice msOffice = new MSOffice(ExcelLocation.SHAREPOINT, applicationId, resourceId, refreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

        List<String> sheets = readExcel.getListOfSheets(msOffice);
        System.out.println(sheets);

        List<List<Object>> data = readExcel.getExcelData(msOffice, "Salesforce", "UsedRange");
        System.out.println(data);
    }

    @Test
    public void testOneDriveExcel() throws IOException {

        String applicationId = "d1c318de-dcb6-4e35-a1ad-15907c7b8744";
        String refreshToken = "<refresh_token>";

        String resourceId = "d1c318de-dcb6-4e35-a1ad-15907c7b8744";

        MSOffice office = new MSOffice(ExcelLocation.ONE_DRIVE,applicationId, resourceId, refreshToken);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

        List<String> sheets = readExcel.getListOfSheets(office);
        System.out.println(sheets);
    }

    @Test
    public void testLocalExcel() throws IOException {
        String localExcelPath = getClass().getClassLoader().getResource("SmokeTests.xlsx").getPath();
        MSOffice office = new MSOffice(ExcelLocation.LOCAL,localExcelPath);
        ReadExcel<MSOffice> readExcel = new ReadExcel<>();

        List<String> sheets = readExcel.getListOfSheets(office);
        System.out.println(sheets);
    }
