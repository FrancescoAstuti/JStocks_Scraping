package afin.jstocks;

import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Scanner;
import org.json.JSONObject;
import org.json.JSONArray;
import java.math.BigDecimal;
import java.math.RoundingMode;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Locale;

public class KeyMetricsTTM {
    private static final String API_KEY = "eb7366217370656d66a56a057b8511b0";
   // private static final String API_URL_RatiosTTM = "https://financialmodelingprep.com/api/v3/ratios-ttm/";
    private static final String API_URL_KeyMetricsTTM = "https://financialmodelingprep.com/api/v3/key-metrics-ttm/";

    public static String getEPSTTM(String ticker) throws IOException {
        String urlString = API_URL_KeyMetricsTTM + ticker + "?apikey=" + API_KEY;
        URL url = new URL(urlString);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.connect();

        int responseCode = conn.getResponseCode();
        
        if (responseCode == 401) {
            throw new RuntimeException("Unauthorized: Invalid API Key");
        } else if (responseCode != 200) {
            throw new RuntimeException("HttpResponseCode: " + responseCode);
        } else {
            String inline = "";
            Scanner scanner = new Scanner(url.openStream());

            // Write all the JSON data into a string using a scanner
            while (scanner.hasNext()) {
                inline += scanner.nextLine();
            }

            // Close the scanner
            scanner.close();

            // Parse the string into a JSON object
            JSONArray jsonArray = new JSONArray(inline);
            if (jsonArray.length() == 0) {
                throw new RuntimeException("No data found for ticker: " + ticker);
            }
            JSONObject jsonObject = jsonArray.getJSONObject(0);

            // Get the EPS TTM from the JSON object
            return jsonObject.getString("epsTTM");
        }
    }

    public static String getROETTM(String ticker) throws IOException {
        String urlString = API_URL_KeyMetricsTTM + ticker + "?apikey=" + API_KEY;
        URL url = new URL(urlString);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.connect();

        int responseCode = conn.getResponseCode();

        if (responseCode == 401) {
            throw new RuntimeException("Unauthorized: Invalid API Key");
        } else if (responseCode != 200) {
            throw new RuntimeException("HttpResponseCode: " + responseCode);
        } else {
            String inline = "";
            Scanner scanner = new Scanner(url.openStream());

            // Write all the JSON data into a string using a scanner
            while (scanner.hasNext()) {
                inline += scanner.nextLine();
            }

            // Close the scanner
            scanner.close();

            // Parse the string into a JSON object
            JSONArray jsonArray = new JSONArray(inline);
            if (jsonArray.length() == 0) {
                throw new RuntimeException("No data found for ticker: " + ticker);
            }
            JSONObject jsonObject = jsonArray.getJSONObject(0);

            // Get the ROE TTM from the JSON object
            return jsonObject.getString("roeTTM");
        }
    }
    
    public static double getDebtToEquityTTM(String ticker) throws IOException {
        String urlString = API_URL_KeyMetricsTTM + ticker + "?apikey=" + API_KEY;
        URL url = new URL(urlString);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.connect();

        int responseCode = conn.getResponseCode();
        
        if (responseCode == 401) {
            throw new RuntimeException("Unauthorized: Invalid API Key");
        } else if (responseCode != 200) {
            throw new RuntimeException("HttpResponseCode: " + responseCode);
        } else {
            String inline = "";
            Scanner scanner = new Scanner(url.openStream());

            // Write all the JSON data into a string using a scanner
            while (scanner.hasNext()) {
                inline += scanner.nextLine();
            }

            // Close the scanner
            scanner.close();

            // Parse the string into a JSON object
            JSONArray jsonArray = new JSONArray(inline);
            if (jsonArray.length() == 0) {
                throw new RuntimeException("No data found for ticker: " + ticker);
            }
            JSONObject jsonObject = jsonArray.getJSONObject(0);

            // Get the dividend yield TTM from the JSON object
            double debtToEquity = jsonObject.getDouble("debtToEquityTTM");

            // Round to one decimal place using BigDecimal
            debtToEquity = new BigDecimal(debtToEquity)
                               .setScale(2, RoundingMode.HALF_UP)
                               .doubleValue();

            //return debtToEquity;
            return 0;
        }
    }
    
    

    public static double getDividendYieldTTM(String ticker) throws IOException {
        String urlString = API_URL_KeyMetricsTTM + ticker + "?apikey=" + API_KEY;
        URL url = new URL(urlString);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.connect();

        int responseCode = conn.getResponseCode();
        
        if (responseCode == 401) {
            throw new RuntimeException("Unauthorized: Invalid API Key");
        } else if (responseCode != 200) {
            throw new RuntimeException("HttpResponseCode: " + responseCode);
        } else {
            String inline = "";
            Scanner scanner = new Scanner(url.openStream());

            // Write all the JSON data into a string using a scanner
            while (scanner.hasNext()) {
                inline += scanner.nextLine();
            }

            // Close the scanner
            scanner.close();

            // Parse the string into a JSON object
            JSONArray jsonArray = new JSONArray(inline);
            if (jsonArray.length() == 0) {
                throw new RuntimeException("No data found for ticker: " + ticker);
            }
            JSONObject jsonObject = jsonArray.getJSONObject(0);

            // Get the dividend yield TTM from the JSON object
            double dividendYieldTTM = jsonObject.getDouble("dividendYieldPercentageTTM");

            // Round to one decimal place using BigDecimal
            dividendYieldTTM = new BigDecimal(dividendYieldTTM)
                               .setScale(1, RoundingMode.HALF_UP)
                               .doubleValue();

            return dividendYieldTTM;
        }
    }
    
    /**
 * Fetches the Current Ratio from Yahoo Finance key-statistics page
 * @param ticker The stock ticker symbol
 * @return The current ratio value as a double, or 0.0 if not found
 */
public static double fetchCurrentRatio(String ticker) {
    try {
        // Format the Yahoo Finance URL for the key-statistics page
        String url = String.format("https://finance.yahoo.com/quote/%s/key-statistics/", ticker);
        
        // Connect to the website with a user agent to mimic a browser
        Document doc = Jsoup.connect(url)
                .userAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
                .timeout(10000)  // 10-second timeout
                .get();
        
        // Find the table row containing Current Ratio
        Elements rows = doc.select("td.label");
        for (Element labelCell : rows) {
            String labelText = labelCell.text().trim();
            
            // Look for "Current Ratio" label - also match variations like "Current Ratio (mrq)"
            if (labelText.contains("Current Ratio")) {
                // Get the next cell which contains the value
                Element valueCell = labelCell.nextElementSibling();
                if (valueCell != null) {
                    String valueText = valueCell.text().trim();
                    System.out.println("Found Current Ratio for " + ticker + ": " + valueText);
                    
                    // Parse the value, handling locale issues
                    return parseDouble(valueText);
                }
            }
        }
        
        // If not found by label text, try using CSS selectors directly based on classes
        Elements ratioLabels = doc.select("td.label.yf-vaowmx");
        for (Element label : ratioLabels) {
            if (label.text().contains("Current Ratio")) {
                Element valueElement = label.nextElementSibling();
                if (valueElement != null) {
                    String valueText = valueElement.text().trim();
                    System.out.println("Found Current Ratio using class selector for " + ticker + ": " + valueText);
                    return parseDouble(valueText);
                }
            }
        }
        
        System.out.println("Current Ratio not found for " + ticker + " on Yahoo Finance");
        return 0.0;
        
    } catch (IOException e) {
        System.err.println("Error fetching Current Ratio for " + ticker + ": " + e.getMessage());
        return 0.0;
    } catch (Exception e) {
        System.err.println("Unexpected error fetching Current Ratio for " + ticker + ": " + e.getMessage());
        e.printStackTrace();
        return 0.0;
    }
}

/**
 * Parses a string to a double, handling different locale decimal separators
 */
private static double parseDouble(String value) {
    if (value == null || value.trim().isEmpty()) {
        return 0.0;
    }
    
    try {
        // Make a copy of the original value for logging
        String originalValue = value;
        System.out.println("Attempting to parse Current Ratio: '" + originalValue + "'");
        
        // Always replace commas with periods to handle both locale formats
        value = value.replace(',', '.');
        
        // Remove any non-numeric characters except decimal point and minus sign
        String cleanValue = value.replaceAll("[^0-9.-]", "");
        System.out.println("Cleaned Current Ratio value for parsing: '" + cleanValue + "'");
        
        // If we ended up with an empty string, return 0
        if (cleanValue.isEmpty()) {
            return 0.0;
        }
        
        // Ensure US locale is used to parse the number
        double result = Double.parseDouble(cleanValue);
        System.out.println("Successfully parsed Current Ratio '" + originalValue + "' to " + result);
        return round(result, 2);
    } catch (NumberFormatException e) {
        System.err.println("Failed to parse Current Ratio value: '" + value + "' - " + e.getMessage());
        
        // Try an alternate approach with explicit locale
        try {
            java.text.NumberFormat format = java.text.NumberFormat.getInstance(Locale.US);
            Number number = format.parse(value.replace(',', '.'));
            return round(number.doubleValue(), 2);
        } catch (Exception ex) {
            System.err.println("Second parsing attempt for Current Ratio also failed: " + ex.getMessage());
            return 0.0;
        }
    }
}

/**
 * Safely rounds a double to the specified number of decimal places
 */
private static double round(double value, int places) {
    if (places < 0) {
        throw new IllegalArgumentException("Cannot round to negative decimal places");
    }
    
    try {
        BigDecimal bd = BigDecimal.valueOf(value);
        bd = bd.setScale(places, RoundingMode.HALF_UP);
        return bd.doubleValue();
    } catch (Exception e) {
        System.err.println("Error rounding Current Ratio value: " + e.getMessage());
        return value; // Return original value if rounding fails
    }
}
}