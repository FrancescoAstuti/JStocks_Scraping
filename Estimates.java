package afin.jstocks;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;

public class Estimates {
    private static final String API_KEY = "BLABLA_eb7366217370656d66a56a057b8511b0";

    public static JSONObject fetchEpsEstimates(String ticker) {
        String urlString = String.format("https://financialmodelingprep.com/api/v3/analyst-estimates/%s?apikey=%s", ticker, API_KEY);
        HttpURLConnection connection = null;

        try {
            URL url = new URL(urlString);
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");

            int responseCode = connection.getResponseCode();
            if (responseCode == 200) {
                Scanner scanner = new Scanner(url.openStream());
                String response = scanner.useDelimiter("\\Z").next();
                scanner.close();

                JSONArray data = new JSONArray(response);
                if (data.length() > 0) {
                    JSONObject result = new JSONObject();

                    // Get current year
                    int currentYear = Calendar.getInstance().get(Calendar.YEAR);

                    // Create a map to store EPS estimates by year
                    Map<Integer, Double> epsEstimatesByYear = new HashMap<>();

                    // Populate the map with available estimates from the API response
                    for (int i = 0; i < data.length(); i++) {
                        JSONObject estimate = data.getJSONObject(i);
                        String date = estimate.getString("date");
                        int year = Integer.parseInt(date.substring(0, 4));  // Extract the year from the date string
                        double eps = estimate.optDouble("estimatedEpsAvg", 0.0);
                        epsEstimatesByYear.put(year, eps);
                    }

                    // Get EPS estimates for the current year and the next four years
                    for (int i = 0; i < 5; i++) {
                        int year = currentYear + i;
                        double eps = epsEstimatesByYear.getOrDefault(year, 0.0);
                        result.put("eps" + i, eps);  // Store EPS estimates with keys "eps0" to "eps4"
                    }

                    return result;
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }
        return null;
    }
}