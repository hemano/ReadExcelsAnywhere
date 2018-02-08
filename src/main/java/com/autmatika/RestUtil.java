package com.autmatika;

import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;
import org.asynchttpclient.AsyncHttpClient;

import java.io.IOException;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;

import static io.restassured.RestAssured.given;
import static org.asynchttpclient.Dsl.asyncHttpClient;

public class RestUtil {

    public Response executeGetRequest(String baseURI, String bearerToken, String path) throws IllegalArgumentException {
        bearerToken = "Bearer " + bearerToken;

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
//        String endPoint = path;
//        if (baseURI.length() != 0) {
//            endPoint = baseURI + "/" + path;
//        }
//
//        org.asynchttpclient.Response whenResponse = null;
//        try (AsyncHttpClient asyncHttpClient = asyncHttpClient()) {
//            whenResponse = asyncHttpClient
//                    .prepareGet(endPoint)
//                    .addHeader("Authorization", bearerToken)
//                    .execute()
//                    .get(60, TimeUnit.SECONDS);
//
//            if(whenResponse.isRedirected()){
//                whenResponse = executeGetRequest("", bearerToken, whenResponse.getHeaders().get("Location"));
//            }
//
//        } catch (InterruptedException e) {
//            e.printStackTrace();
//        } catch (ExecutionException e) {
//            e.printStackTrace();
//        } catch (TimeoutException e) {
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//
//
//        return whenResponse;

        return response;
    }


    /**
     * @return Returns the response object of RestAssured Api
     */
    public org.asynchttpclient.Response executePostRequest(String baseURI, String bearerToken, String requestBody) throws IllegalArgumentException, IOException {

        System.out.println("REQUEST: \n" + requestBody);

        bearerToken = "Bearer " + bearerToken;

        org.asynchttpclient.Response whenResponse = null;
        try (AsyncHttpClient asyncHttpClient = asyncHttpClient()) {
            whenResponse = asyncHttpClient
                    .preparePost(baseURI)
                    .addHeader("Content-Type", "application/json")
                    .addHeader("Authorization", bearerToken)
                    .setBody(requestBody)
                    .execute()
                    .get(50, TimeUnit.SECONDS);
        } catch (InterruptedException e) {
            e.printStackTrace();
        } catch (ExecutionException e) {
            e.printStackTrace();
        } catch (TimeoutException e) {
            e.printStackTrace();
        }

//        Response response = given()
//                .contentType("application/json")
//                .header("Authorization", bearerToken)
//                .baseUri(baseURI)
//                .when()
//                .body(requestBody)
//                .post();
//
//        System.out.println(response.getTime());
//
//        String responseString = response.asString();
//        JsonPath jsonPath = new JsonPath(responseString);
//
//        if (response.getStatusCode() != 200) {
//            throw new IllegalArgumentException("Error: " + jsonPath.get("error") + " Error description: " + jsonPath.get("error_description"));
//        }

        if (whenResponse.getStatusCode() != 200) {
            throw new IllegalArgumentException("Error");
        }

        return whenResponse;

    }

}