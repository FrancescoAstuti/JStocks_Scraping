package afin.jstocks;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

public class Ratios {

    private static final String API_KEY = API_key.getApiKey();

    public static List<RatioData> fetchHistoricalPE(String ticker) {
        List<RatioData> peRatios = new ArrayList<>();
        try {
            URL url = new URL("https://financialmodelingprep.com/api/v3/ratios/" + ticker + "?period=annual&apikey=" + API_KEY);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");

            int responseCode = conn.getResponseCode();
            System.out.println("PE API Response Code: " + responseCode);  // Debugging statement

            BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
            String inputLine;
            StringBuilder content = new StringBuilder();
            while ((inputLine = in.readLine()) != null) {
                content.append(inputLine);
            }

            in.close();
            conn.disconnect();

            System.out.println("PE API Response: " + content.toString());  // Debugging statement

            JSONArray jsonArray = new JSONArray(content.toString());
            if (jsonArray.length() == 0) {
                System.out.println("No historical PE data available for ticker: " + ticker);
            } else {
                for (int i = 0; i < jsonArray.length(); i++) {
                    JSONObject jsonObject = jsonArray.getJSONObject(i);
                    String date = jsonObject.getString("date");
                    double pe = jsonObject.getDouble("priceEarningsRatio");
                    peRatios.add(new RatioData(date, pe));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return peRatios;
    }
    
    

    public static List<RatioData> fetchHistoricalPB(String ticker) {
        List<RatioData> pbRatios = new ArrayList<>();
        try {
            URL url = new URL("https://financialmodelingprep.com/api/v3/ratios/" + ticker + "?period=annual&apikey=" + API_KEY);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");

            int responseCode = conn.getResponseCode();
            System.out.println("PB API Response Code: " + responseCode);  // Debugging statement

            BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
            String inputLine;
            StringBuilder content = new StringBuilder();
            while ((inputLine = in.readLine()) != null) {
                content.append(inputLine);
            }

            in.close();
            conn.disconnect();

            System.out.println("PB API Response: " + content.toString());  // Debugging statement

            JSONArray jsonArray = new JSONArray(content.toString());
            if (jsonArray.length() == 0) {
                System.out.println("No historical PB data available for ticker: " + ticker);
            } else {
                for (int i = 0; i < jsonArray.length(); i++) {
                    JSONObject jsonObject = jsonArray.getJSONObject(i);
                    String date = jsonObject.getString("date");
                    double pb = jsonObject.getDouble("priceToBookRatio");
                    pbRatios.add(new RatioData(date, pb));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return pbRatios;
    }

    public static List<RatioData> fetchQuarterlyEPS(String ticker) {
        List<RatioData> epsRatios = new ArrayList<>();
        try {
            URL url = new URL("https://financialmodelingprep.com/api/v3/income-statement/" + ticker + "?period=annual&apikey=" + API_KEY);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");

            int responseCode = conn.getResponseCode();
            System.out.println("EPS API Response Code: " + responseCode);  // Debugging statement

            BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
            String inputLine;
            StringBuilder content = new StringBuilder();
            while ((inputLine = in.readLine()) != null) {
                content.append(inputLine);
            }

            in.close();
            conn.disconnect();

            System.out.println("EPS API Response: " + content.toString());  // Debugging statement

            JSONArray jsonArray = new JSONArray(content.toString());
            if (jsonArray.length() == 0) {
                System.out.println("No quarterly EPS data available for ticker: " + ticker);
            } else {
                for (int i = 0; i < jsonArray.length(); i++) {
                    JSONObject jsonObject = jsonArray.getJSONObject(i);
                    String date = jsonObject.getString("date");
                    double eps = jsonObject.getDouble("eps");
                    epsRatios.add(new RatioData(date, eps));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return epsRatios;
    }

    // New method to fetch historical priceToFreeCashFlowsRatio
    public static List<RatioData> fetchHistoricalPFCF(String ticker) {
        List<RatioData> pfcfRatios = new ArrayList<>();
        try {
            URL url = new URL("https://financialmodelingprep.com/api/v3/ratios/" + ticker + "?period=annual&apikey=" + API_KEY);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");

            int responseCode = conn.getResponseCode();
            System.out.println("PFCF API Response Code: " + responseCode);  // Debugging statement

            BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
            String inputLine;
            StringBuilder content = new StringBuilder();
            while ((inputLine = in.readLine()) != null) {
                content.append(inputLine);
            }

            in.close();
            conn.disconnect();

            System.out.println("PFCF API Response: " + content.toString());  // Debugging statement

            JSONArray jsonArray = new JSONArray(content.toString());
            if (jsonArray.length() == 0) {
                System.out.println("No historical PFCF data available for ticker: " + ticker);
            } else {
                for (int i = 0; i < jsonArray.length(); i++) {
                    JSONObject jsonObject = jsonArray.getJSONObject(i);
                    String date = jsonObject.getString("date");
                    double pfcf = jsonObject.getDouble("priceToFreeCashFlowsRatio");
                    pfcfRatios.add(new RatioData(date, pfcf));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return pfcfRatios;
    }

    // New method to fetch historical debtToEquityRatio
    public static List<RatioData> fetchHistoricalDebtToEquity(String ticker) {
        List<RatioData> debtToEquityRatios = new ArrayList<>();
        try {
            URL url = new URL("https://financialmodelingprep.com/api/v3/ratios/" + ticker + "?period=annual&apikey=" + API_KEY);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");

            int responseCode = conn.getResponseCode();
            System.out.println("Debt to Equity API Response Code: " + responseCode);  // Debugging statement

            BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
            String inputLine;
            StringBuilder content = new StringBuilder();
            while ((inputLine = in.readLine()) != null) {
                content.append(inputLine);
            }

            in.close();
            conn.disconnect();

            System.out.println("Debt to Equity API Response: " + content.toString());  // Debugging statement

            JSONArray jsonArray = new JSONArray(content.toString());
            if (jsonArray.length() == 0) {
                System.out.println("No historical Debt to Equity data available for ticker: " + ticker);
            } else {
                for (int i = 0; i < jsonArray.length(); i++) {
                    JSONObject jsonObject = jsonArray.getJSONObject(i);
                    String date = jsonObject.getString("date");
                    double debtToEquity = jsonObject.getDouble("debtEquityRatio");
                    debtToEquityRatios.add(new RatioData(date, debtToEquity));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return debtToEquityRatios;
    }
    
    public static List<RatioData> fetchHistoricalReturnOnEquity(String ticker) {
        List<RatioData> ReturnOnEquityRatios = new ArrayList<>();
        try {
            URL url = new URL("https://financialmodelingprep.com/api/v3/ratios/" + ticker + "?period=annual&apikey=" + API_KEY);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");

            int responseCode = conn.getResponseCode();
            
            BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
            String inputLine;
            StringBuilder content = new StringBuilder();
            while ((inputLine = in.readLine()) != null) {
                content.append(inputLine);
            }

            in.close();
            conn.disconnect();
           
            JSONArray jsonArray = new JSONArray(content.toString());
            if (jsonArray.length() == 0) {
                System.out.println("No historical Return on Equity data available for ticker: " + ticker);
            } else {
                for (int i = 0; i < jsonArray.length(); i++) {
                    JSONObject jsonObject = jsonArray.getJSONObject(i);
                    String date = jsonObject.getString("date");
                    double returnOnEquity = jsonObject.getDouble("returnOnEquity");
                    ReturnOnEquityRatios.add(new RatioData(date, returnOnEquity));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return ReturnOnEquityRatios;
    }
    
    // New method to fetch historical payout ratios from local files
    public static List<RatioData> fetchHistoricalPayoutRatio(String ticker) {
        List<RatioData> payoutRatios = new ArrayList<>();
        File file = new File("company_data/" + ticker + "_Financial_Ratios_FY.json");
        
        if (!file.exists()) {
            System.out.println("No financial ratios file found for ticker: " + ticker);
            return payoutRatios;
        }
        
        try (BufferedReader reader = new BufferedReader(new FileReader(file))) {
            StringBuilder content = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                content.append(line);
            }
            
            JSONArray jsonArray = new JSONArray(content.toString());
            if (jsonArray.length() == 0) {
                System.out.println("No historical payout ratio data available for ticker: " + ticker);
            } else {
                for (int i = 0; i < jsonArray.length(); i++) {
                    JSONObject jsonObject = jsonArray.getJSONObject(i);
                    String date = jsonObject.getString("date");
                    double payoutRatio = jsonObject.optDouble("payoutRatio", 0.0);
                    payoutRatios.add(new RatioData(date, payoutRatio));
                }
            }
        } catch (Exception e) {
            System.err.println("Error reading payout ratio data for " + ticker + ": " + e.getMessage());
            e.printStackTrace();
        }
        
        return payoutRatios;
    }
    
    // Method to calculate average payout ratio
    public static double fetchPayoutRatioAverage(String ticker) {
        List<RatioData> payoutRatios = fetchHistoricalPayoutRatio(ticker);
        
        if (payoutRatios.isEmpty()) {
            return 0.0;
        }
        
        double sum = 0;
        int count = 0;
        
        for (RatioData ratio : payoutRatios) {
            double value = ratio.getValue();
            if (value >= 0) {  // Only include non-negative payout ratios
                // Cap extreme values to prevent outliers from skewing the average
                value = Math.min(value, 1.0);  // Cap at 100% (1.0)
                sum += value;
                count++;
            }
        }
        
        return count > 0 ? round(sum / count, 2) : 0.0;
    }
    
    public static double fetchDebtToEquityAverage(String ticker) {
        List<RatioData> debtToEquityRatios = fetchHistoricalDebtToEquity(ticker);
        
        if (debtToEquityRatios.isEmpty()) {
            return 0.0;
        }

        double sum = 0;
        int count = 0;
        
        for (RatioData ratio : debtToEquityRatios) {
            double value = ratio.getValue();
            if (value != 0) {
                value = Math.min(Math.max(value, -3.0), 3.0); // Cap between -3 and 3
                sum += value;
                count++;
            }
        }
        
        return count > 0 ? round(sum / count, 2) : 0.0;
    }
    
    public static double fetchPriceToFreeCashFlowAverage(String ticker) {
        List<RatioData> pfcfRatios = fetchHistoricalPFCF(ticker);
        
        if (pfcfRatios.isEmpty()) {
            return 0.0;
        }

        double sum = 0;
        int count = 0;
        
        for (RatioData ratio : pfcfRatios) {
            double value = ratio.getValue();
            if (value != 0) {
                // Cap extreme values to prevent outliers from skewing the average
                value = Math.min(Math.max(value, -100.0), 100.0); // Cap between -100 and 100
                sum += value;
                count++;
            }
        }
        
        return count > 0 ? round(sum / count, 2) : 0.0;
    } 
    
    private static double round(double value, int places) {
        if (places < 0) throw new IllegalArgumentException();
        long factor = (long) Math.pow(10, places);
        value = value * factor;
        long tmp = Math.round(value);
        return (double) tmp / factor;
    }
}

class RatioData {
    private String date;
    private double value;

    public RatioData(String date, double value) {
        this.date = date;
        this.value = value;
    }

    public String getDate() {
        return date;
    }

    public double getValue() {
        return value;
    }
}