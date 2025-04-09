package afin.jstocks;

import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.util.Calendar;

public class Estimates {
    // Remove the API key since we're not using the API anymore
    
       
    public static JSONObject fetchEpsEstimates(String ticker) {
        // Keep the original ticker format for Yahoo Finance (don't convert .L to .LON)
        String url = "https://finance.yahoo.com/quote/" + ticker + "/analysis?p=" + ticker;
        
        try {
            // Connect to Yahoo Finance with a user agent to avoid blocking
            Document doc = Jsoup.connect(url)
                    .userAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
                    .timeout(10000)
                    .get();

            // Target the Earnings Estimate table
            Elements rows = doc.select("tr:has(td:contains(Avg. Estimate))");
            
            if (!rows.isEmpty()) {
                Element avgEstimateRow = rows.first();
                Elements cells = avgEstimateRow.select("td");
                
                // Create the result JSON object
                JSONObject result = new JSONObject();
                
                // Get current year
                int currentYear = Calendar.getInstance().get(Calendar.YEAR);
                
                // Parse the values for current year and next years
                // Yahoo provides current quarter, next quarter, current year, next year
                
                if (cells.size() >= 4) {  // Ensure we have enough cells
                    double epsCurrentYear = parseEps(cells.get(3).text());
                    result.put("eps0", epsCurrentYear);
                    
                    if (cells.size() >= 5) {
                        double epsNextYear = parseEps(cells.get(4).text());
                        result.put("eps1", epsNextYear);
                        
                        // We don't have data for years beyond next year, so estimate based on growth rate
                        if (epsCurrentYear > 0 && epsNextYear > 0) {
                            double growthRate = (epsNextYear - epsCurrentYear) / epsCurrentYear;
                            double epsYear2 = epsNextYear * (1 + growthRate);
                            double epsYear3 = epsYear2 * (1 + growthRate);
                            double epsYear4 = epsYear3 * (1 + growthRate);
                            
                            result.put("eps2", round(epsYear2, 2));
                            result.put("eps3", round(epsYear3, 2));
                            result.put("eps4", round(epsYear4, 2));
                        } else {
                            result.put("eps2", 0.0);
                            result.put("eps3", 0.0);
                            result.put("eps4", 0.0);
                        }
                    } else {
                        result.put("eps1", 0.0);
                        result.put("eps2", 0.0);
                        result.put("eps3", 0.0);
                        result.put("eps4", 0.0);
                    }
                } else {
                    // Not enough data
                    result.put("eps0", 0.0);
                    result.put("eps1", 0.0);
                    result.put("eps2", 0.0);
                    result.put("eps3", 0.0);
                    result.put("eps4", 0.0);
                }
                
                return result;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        
        // Return a default object with zeros if we couldn't get the data
        JSONObject defaultResult = new JSONObject();
        defaultResult.put("eps0", 0.0);
        defaultResult.put("eps1", 0.0);
        defaultResult.put("eps2", 0.0);
        defaultResult.put("eps3", 0.0);
        defaultResult.put("eps4", 0.0);
        return defaultResult;
    }
    
    private static double parseEps(String epsText) {
        try {
            // Remove any non-numeric characters except decimal point and negative sign
            epsText = epsText.replaceAll("[^0-9.-]", "");
            return Double.parseDouble(epsText);
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }
    
    private static double round(double value, int places) {
        if (places < 0) throw new IllegalArgumentException();
        
        long factor = (long) Math.pow(10, places);
        value = value * factor;
        long tmp = Math.round(value);
        return (double) tmp / factor;
    }
}