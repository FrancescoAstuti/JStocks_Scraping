package afin.jstocks;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;

public class API_key {
    private static final String API_KEY_FILE_PATH = "api_key.txt";

    public static String getApiKey() {
        String apiKey = "";
        try (BufferedReader reader = new BufferedReader(new FileReader(API_KEY_FILE_PATH))) {
            apiKey = reader.readLine().trim();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return apiKey;
    }
}