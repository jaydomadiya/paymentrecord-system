package com.paymentrecord.config;

import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;

import java.io.FileInputStream;
import java.util.Collections;

public class GoogleSheetConfig {


    public static Sheets getSheetsService() throws Exception {

        GoogleCredentials credentials = GoogleCredentials
                .fromStream(new FileInputStream("src/main/resources/cred.json"))
                .createScoped(Collections.singleton("https://www.googleapis.com/auth/spreadsheets"));

        return new Sheets.Builder(
                GoogleNetHttpTransport.newTrustedTransport(),
                GsonFactory.getDefaultInstance(),
                new HttpCredentialsAdapter(credentials)
        )
                .setApplicationName("Payment Dashboard")
                .build();
    }
}