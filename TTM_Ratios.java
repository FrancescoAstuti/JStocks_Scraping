package afin.jstocks;

import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Scanner;
import org.json.JSONArray;
import org.json.JSONObject;
import java.math.BigDecimal;
import java.math.RoundingMode;

public class TTM_Ratios {
    private static final String API_KEY = API_key.getApiKey();
    
    public static double getPriceToFreeCashFlowRatioTTM(String ticker) {
        String urlString = String.format("https://financialmodelingprep.com/api/v3/ratios-ttm/%s?apikey=%s", ticker, API_KEY);
        HttpURLConnection connection = null;
        
        try {
            URL url = new URL(urlString);
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(5000);
            
            int responseCode = connection.getResponseCode();
            if (responseCode == 200) {
                Scanner scanner = new Scanner(url.openStream());
                String response = scanner.useDelimiter("\\Z").next();
                scanner.close();
                
                JSONArray data = new JSONArray(response);
                if (data.length() > 0) {
                    JSONObject ratios = data.getJSONObject(0);
                    double ratio = ratios.optDouble("priceToFreeCashFlowsRatioTTM", 0.0);
                    return round(ratio, 2);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("Error fetching P/FCF for " + ticker + ": " + e.getMessage());
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }
        
        return 0.0;
    }
    
    private static double round(double value, int places) {
        if (places < 0) throw new IllegalArgumentException();
        BigDecimal bd = BigDecimal.valueOf(value);
        bd = bd.setScale(places, RoundingMode.HALF_UP);
        return bd.doubleValue();
    }
}