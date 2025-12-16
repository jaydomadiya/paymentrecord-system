//package com.paymentrecord.config;
//
//import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
//import com.google.api.client.json.gson.GsonFactory;
//import com.google.api.services.sheets.v4.Sheets;
//import com.google.auth.http.HttpCredentialsAdapter;
//import com.google.auth.oauth2.GoogleCredentials;
//
//import java.io.FileInputStream;
//import java.util.Collections;
//
//public class GoogleSheetConfig {
//
//
//    public static Sheets getSheetsService() throws Exception {
//
//        GoogleCredentials credentials = GoogleCredentials
//                .fromStream(new FileInputStream("src/main/resources/cred.json"))
//                .createScoped(Collections.singleton("https://www.googleapis.com/auth/spreadsheets"));
//
//        return new Sheets.Builder(
//                GoogleNetHttpTransport.newTrustedTransport(),
//                GsonFactory.getDefaultInstance(),
//                new HttpCredentialsAdapter(credentials)
//        )
//                .setApplicationName("Payment Dashboard")
//                .build();
//    }
//}

package com.paymentrecord.config;

import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;

import java.io.ByteArrayInputStream;
import java.util.Base64;
import java.util.Collections;

public class GoogleSheetConfig {

    public static Sheets getSheetsService() throws Exception {

        // 1️⃣ Render ENV variable read kare
        String base64Cred = System.getenv("GOOGLE_CREDENTIALS");

        if (base64Cred == null || base64Cred.isEmpty()) {
            throw new RuntimeException("GOOGLE_CREDENTIALS not set in Render");
        }

        // 2️⃣ Base64 decode
        byte[] decoded = Base64.getDecoder().decode(base64Cred);

        // 3️⃣ JSON stream banave (memory ma)
        ByteArrayInputStream stream = new ByteArrayInputStream(decoded);

        // 4️⃣ Google credential object
        GoogleCredentials credentials = GoogleCredentials
                .fromStream(stream)
                .createScoped(Collections.singleton(
                        "https://www.googleapis.com/auth/spreadsheets"
                ));

        // 5️⃣ Sheets service return kare
        return new Sheets.Builder(
                GoogleNetHttpTransport.newTrustedTransport(),
                GsonFactory.getDefaultInstance(),
                new HttpCredentialsAdapter(credentials)
        )
                .setApplicationName("Payment Dashboard")
                .build();
    }
}
