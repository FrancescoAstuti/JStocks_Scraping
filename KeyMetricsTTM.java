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
import java.util.Locale;
import java.util.logging.Logger;
import java.util.logging.Level;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;

public class KeyMetricsTTM {
    private static final Logger LOGGER = Logger.getLogger(KeyMetricsTTM.class.getName());
    private static final String API_KEY = "eb7366217370656d66a56a057b8511b0";
    private static final String API_URL_KEY_METRICS_TTM = "https://financialmodelingprep.com/api/v3/key-metrics-ttm/";
    private static final String YAHOO_FINANCE_STATS_URL = "https://finance.yahoo.com/quote/%s/key-statistics/";
    private static final String USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
    private static final int TIMEOUT = 10000; // 10 seconds
    
    // Decimal formatter with comma as decimal separator
    private static final DecimalFormat DECIMAL_FORMATTER;
    
    static {
        DecimalFormatSymbols symbols = new DecimalFormatSymbols(Locale.getDefault());
        symbols.setDecimalSeparator(',');
        DECIMAL_FORMATTER = new DecimalFormat("0.00", symbols);
        DECIMAL_FORMATTER.setRoundingMode(RoundingMode.HALF_UP);
    }

    /**
     * Fetches a specific metric from Financial Modeling Prep API
     * 
     * @param ticker The stock ticker symbol
     * @param metricName The name of the metric to fetch
     * @return The value of the specified metric as a String
     * @throws IOException If there's an error connecting to the API
     */
    private static String getMetricFromFMP(String ticker, String metricName) throws IOException {
        String urlString = API_URL_KEY_METRICS_TTM + ticker + "?apikey=" + API_KEY;
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
            StringBuilder inline = new StringBuilder();
            try (Scanner scanner = new Scanner(url.openStream())) {
                while (scanner.hasNext()) {
                    inline.append(scanner.nextLine());
                }
            }

            JSONArray jsonArray = new JSONArray(inline.toString());
            if (jsonArray.length() == 0) {
                throw new RuntimeException("No data found for ticker: " + ticker);
            }
            JSONObject jsonObject = jsonArray.getJSONObject(0);

            return jsonObject.getString(metricName);
        }
    }

    /**
     * Gets EPS TTM from Financial Modeling Prep API
     * 
     * @param ticker The stock ticker symbol
     * @return The EPS TTM value as a String
     * @throws IOException If there's an error connecting to the API
     */
    public static String getEPSTTM(String ticker) throws IOException {
        return getMetricFromFMP(ticker, "epsTTM");
    }

    /**
     * Gets ROE TTM from Financial Modeling Prep API
     * 
     * @param ticker The stock ticker symbol
     * @return The ROE TTM value as a String
     * @throws IOException If there's an error connecting to the API
     */
    public static String getROETTM(String ticker) throws IOException {
        return getMetricFromFMP(ticker, "roeTTM");
    }

    /**
     * Fetches and connects to Yahoo Finance statistics page for a given ticker
     * 
     * @param ticker The stock ticker symbol
     * @return The Jsoup Document object containing the page content
     * @throws IOException If there's an error connecting to the website
     */
    private static Document getYahooFinanceDocument(String ticker) throws IOException {
        String url = String.format(YAHOO_FINANCE_STATS_URL, ticker);
        return Jsoup.connect(url)
                .userAgent(USER_AGENT)
                .timeout(TIMEOUT)
                .get();
    }

    /**
     * Helper method to extract a financial ratio from Yahoo Finance
     * 
     * @param doc The Jsoup Document containing the Yahoo Finance page
     * @param ticker The stock ticker symbol (for logging)
     * @param ratioLabel The label text to search for (e.g., "Total Debt/Equity")
     * @param convertToDecimal Whether to convert percentage values to decimal form (34% → 0,34)
     * @return The extracted ratio as a double, or 0.0 if not found
     */
    private static double extractFinancialRatio(Document doc, String ticker, String ratioLabel, boolean convertToDecimal) {
        try {
            // Try finding by direct label text first
            Elements rows = doc.select("td.label");
            for (Element labelCell : rows) {
                if (labelCell.text().trim().contains(ratioLabel)) {
                    Element valueCell = labelCell.nextElementSibling();
                    if (valueCell != null) {
                        String valueText = valueCell.text().trim();
                        LOGGER.info("Found " + ratioLabel + " for " + ticker + ": " + valueText);
                        return parseFinancialValue(valueText, convertToDecimal);
                    }
                }
            }
            
            // Try using specific CSS class selector
            Elements ratioLabels = doc.select("td.label.yf-vaowmx");
            for (Element label : ratioLabels) {
                if (label.text().contains(ratioLabel)) {
                    Element valueElement = label.nextElementSibling();
                    if (valueElement != null) {
                        String valueText = valueElement.text().trim();
                        LOGGER.info("Found " + ratioLabel + " using class selector for " + ticker + ": " + valueText);
                        return parseFinancialValue(valueText, convertToDecimal);
                    }
                }
            }
            
            // Try a more specific selector for certain ratios
            if (ratioLabel.equals("Payout Ratio")) {
                Elements specificElements = doc.select("tr.row.yf-vaowmx td.label.yf-vaowmx:contains(" + ratioLabel + ")");
                for (Element label : specificElements) {
                    Element valueElement = label.nextElementSibling();
                    if (valueElement != null && valueElement.hasClass("value") && valueElement.hasClass("yf-vaowmx")) {
                        String valueText = valueElement.text().trim();
                        LOGGER.info("Found " + ratioLabel + " using specific selector for " + ticker + ": " + valueText);
                        return parseFinancialValue(valueText, convertToDecimal);
                    }
                }
            }
            
            LOGGER.warning(ratioLabel + " not found for " + ticker + " on Yahoo Finance");
            return 0.0;
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "Error extracting " + ratioLabel + " for " + ticker, e);
            return 0.0;
        }
    }

    /**
     * Parses a financial value text to a double
     * 
     * @param valueText The text value to parse
     * @param convertToDecimal Whether to convert percentage values to decimal form (34% → 0,34)
     * @return The parsed and processed value
     */
    private static double parseFinancialValue(String valueText, boolean convertToDecimal) {
        boolean isPercentage = valueText.endsWith("%");
        
        if (isPercentage) {
            valueText = valueText.substring(0, valueText.length() - 1);
            double value = parseDouble(valueText);
            
            if (convertToDecimal) {
                value = value / 100.0; // Convert to decimal form
            }
            
            return round(value, 2);
        } else {
            return round(parseDouble(valueText), 2);
        }
    }

    /**
     * Fetches the Debt to Equity ratio from Yahoo Finance
     * 
     * @param ticker The stock ticker symbol
     * @return The debt to equity ratio value as a double, or 0.0 if not found
     */
    public static double fetchDebtToEquity(String ticker) {
        try {
            Document doc = getYahooFinanceDocument(ticker);
            // For Debt/Equity, we keep as is (not converting to decimal)
            return extractFinancialRatio(doc, ticker, "Total Debt/Equity", false);
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "Error fetching Debt/Equity for " + ticker, e);
            return 0.0;
        }
    }

    /**
     * Fetches the Dividend Yield from Yahoo Finance and returns as decimal
     * 
     * @param ticker The stock ticker symbol
     * @return The dividend yield value as a decimal (e.g., 0.0434 for 4.34%), or 0.0 if not found
     */
    public static double fetchDividendYield(String ticker) {
        try {
            Document doc = getYahooFinanceDocument(ticker);
            
            // Try Forward Annual Dividend Yield first (converting to decimal)
            double forwardYield = extractFinancialRatio(doc, ticker, "Forward Annual Dividend Yield", true);
            if (forwardYield > 0.0) {
                return forwardYield;
            }
            
            // Fall back to Trailing Annual Dividend Yield (converting to decimal)
            return extractFinancialRatio(doc, ticker, "Trailing Annual Dividend Yield", true);
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "Error fetching Dividend Yield for " + ticker, e);
            return 0.0;
        }
    }

    /**
     * Fetches the formatted Dividend Yield for display
     * 
     * @param ticker The stock ticker symbol
     * @return The dividend yield value formatted as "0,xx" or empty string if not found
     */
    public static String fetchFormattedDividendYield(String ticker) {
        double yield = fetchDividendYield(ticker);
        if (yield <= 0.0) {
            return "";
        }
        return DECIMAL_FORMATTER.format(yield);
    }

    /**
     * Fetches the Payout Ratio from Yahoo Finance and returns as decimal
     * 
     * @param ticker The stock ticker symbol
     * @return The payout ratio as a decimal (e.g., 0.3474 for 34.74%), or 0.0 if not found
     */
    public static double fetchPayoutRatio(String ticker) {
        try {
            Document doc = getYahooFinanceDocument(ticker);
            return extractFinancialRatio(doc, ticker, "Payout Ratio", true);
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "Error fetching Payout Ratio for " + ticker, e);
            return 0.0;
        }
    }

    /**
     * Fetches the formatted Payout Ratio for display
     * 
     * @param ticker The stock ticker symbol
     * @return The payout ratio formatted as "0,xx" or empty string if not found
     */
    public static String fetchFormattedPayoutRatio(String ticker) {
        double ratio = fetchPayoutRatio(ticker);
        if (ratio <= 0.0) {
            return "";
        }
        return DECIMAL_FORMATTER.format(ratio);
    }

    /**
     * Fetches the Current Ratio from Yahoo Finance
     * 
     * @param ticker The stock ticker symbol
     * @return The current ratio value as a double, or 0.0 if not found
     */
    public static double fetchCurrentRatio(String ticker) {
        try {
            Document doc = getYahooFinanceDocument(ticker);
            return extractFinancialRatio(doc, ticker, "Current Ratio", false);
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "Error fetching Current Ratio for " + ticker, e);
            return 0.0;
        }
    }

    /**
     * Helper method to parse a double value from a string, handling different locale formats
     *
     * @param value The string value to parse
     * @return The parsed double value
     */
    private static double parseDouble(String value) {
        if (value == null || value.trim().isEmpty()) {
            return 0.0;
        }
        
        try {
            // Replace commas with periods to handle different locale formats
            value = value.replace(',', '.');
            
            // Remove any non-numeric characters except decimal point and minus sign
            String cleanValue = value.replaceAll("[^0-9.-]", "");
            
            // If empty after cleaning, return 0
            if (cleanValue.isEmpty()) {
                return 0.0;
            }
            
            return Double.parseDouble(cleanValue);
        } catch (NumberFormatException e) {
            LOGGER.warning("Failed to parse value: '" + value + "' - " + e.getMessage());
            
            // Try alternate parsing with explicit locale
            try {
                java.text.NumberFormat format = java.text.NumberFormat.getInstance(Locale.US);
                Number number = format.parse(value.replace(',', '.'));
                return number.doubleValue();
            } catch (Exception ex) {
                LOGGER.severe("Second parsing attempt also failed: " + ex.getMessage());
                return 0.0;
            }
        }
    }

    /**
     * Helper method to round a double value to a specified number of decimal places
     *
     * @param value The value to round
     * @param places The number of decimal places
     * @return The rounded value
     */
    private static double round(double value, int places) {
        if (places < 0) {
            throw new IllegalArgumentException("Decimal places must be >= 0");
        }
        
        try {
            BigDecimal bd = BigDecimal.valueOf(value);
            bd = bd.setScale(places, RoundingMode.HALF_UP);
            return bd.doubleValue();
        } catch (Exception e) {
            LOGGER.warning("Error rounding value: " + e.getMessage());
            return value; // Return original value if rounding fails
        }
    }
}