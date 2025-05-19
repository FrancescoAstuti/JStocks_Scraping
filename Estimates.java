package afin.jstocks;

import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.IOException;
import java.util.Calendar;

public class Estimates {

    public static JSONObject fetchEpsEstimates(String ticker) {
        return fetchYahooEpsEstimates(ticker);
    }
    
    public static JSONObject fetchYahooEpsEstimates(String ticker) {
        try {
            // Format the Yahoo Finance URL for the analysis page
            String url = String.format("https://finance.yahoo.com/quote/%s/analysis", ticker);
            
            // Connect to the website with a user agent to mimic a browser
            Document doc = Jsoup.connect(url)
                    .userAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
                    .timeout(10000)  // 10-second timeout
                    .get();
            
            // Create a JSON object to store the results
            JSONObject result = new JSONObject();
            
            // More robust table selection - look for any table after Earnings Estimate header
            Element earningsTable = null;
            Elements headers = doc.select("h3, section h2");
            for (Element header : headers) {
                if (header.text().contains("Earnings Estimate")) {
                    // Find the first table after this header
                    Element section = header.parent().parent();
                    Elements tables = section.select("table");
                    if (!tables.isEmpty()) {
                        earningsTable = tables.first();
                        break;
                    }
                }
            }
            
            // If we didn't find the table with the first approach, try alternative selectors
            if (earningsTable == null) {
                // Try direct approach to find the table
                Elements tables = doc.select("table");
                
                // Look for a table that has the expected columns
                for (Element table : tables) {
                    Elements tableHeaders = table.select("thead th");
                    boolean hasCurrentYear = false;
                    boolean hasNextYear = false;
                    
                    for (Element th : tableHeaders) {
                        String headerText = th.text();
                        if (headerText.contains("Current Year")) {
                            hasCurrentYear = true;
                        } else if (headerText.contains("Next Year")) {
                            hasNextYear = true;
                        }
                    }
                    
                    if (hasCurrentYear && hasNextYear) {
                        earningsTable = table;
                        break;
                    }
                }
            }
            
            if (earningsTable != null) {
                // Now find the row with "Avg. Estimate"
                Elements rows = earningsTable.select("tbody tr");
                Element avgEstimateRow = null;
                
                for (Element row : rows) {
                    Elements cells = row.select("td");
                    if (cells.isEmpty()) continue;
                    
                    String firstCellText = cells.first().text().trim();
                    if (firstCellText.contains("Avg. Estimate") || 
                        firstCellText.contains("Average Estimate") ||
                        firstCellText.contains("Avg Estimate")) {
                        avgEstimateRow = row;
                        break;
                    }
                }
                
                // If we still can't find it, take the second row (which typically has the estimates)
                if (avgEstimateRow == null && rows.size() >= 2) {
                    avgEstimateRow = rows.get(1);  // Second row is usually Avg. Estimate
                    System.out.println("Using second row as fallback for Avg. Estimate in " + ticker);
                }
                
                if (avgEstimateRow != null) {
                    Elements cells = avgEstimateRow.select("td");
                    
                    // Print out all cells for debugging
                    System.out.println("Found " + cells.size() + " cells in Avg. Estimate row");
                    for (int i = 0; i < cells.size(); i++) {
                        System.out.println("Cell " + i + ": " + cells.get(i).text());
                    }
                    
                    // Get the current year and next year from the table headers
                    Elements tableHeaders = earningsTable.select("thead th");
                    String currentYearText = "";
                    String nextYearText = "";
                    
                    // Print all headers for debugging
                    System.out.println("Found " + tableHeaders.size() + " headers in table");
                    for (int i = 0; i < tableHeaders.size(); i++) {
                        String headerText = tableHeaders.get(i).text();
                        System.out.println("Header " + i + ": " + headerText);
                        
                        if (headerText.contains("Current Year")) {
                            currentYearText = headerText;
                        } else if (headerText.contains("Next Year")) {
                            nextYearText = headerText;
                        }
                    }
                    
                    int currentYearIndex = -1;
                    int nextYearIndex = -1;
                    
                    // Find indices for Current Year and Next Year columns
                    for (int i = 0; i < tableHeaders.size(); i++) {
                        String headerText = tableHeaders.get(i).text();
                        if (headerText.contains("Current Year")) {
                            currentYearIndex = i;
                        } else if (headerText.contains("Next Year")) {
                            nextYearIndex = i;
                        }
                    }
                    
                    // Fallback to fixed positions if we couldn't find the columns
                    if (currentYearIndex == -1) currentYearIndex = 3;  // Usually 4th column (0-indexed)
                    if (nextYearIndex == -1) nextYearIndex = 4;       // Usually 5th column
                    
                    // Extract years from header text (e.g., "Current Year (2025)" -> "2025")
                    int currentYear = extractYear(currentYearText);
                    int nextYear = extractYear(nextYearText);
                    
                    // If we couldn't extract years from headers, use the current calendar year
                    if (currentYear == 0) {
                        currentYear = Calendar.getInstance().get(Calendar.YEAR);
                        nextYear = currentYear + 1;
                    }
                    
                    // Get the estimates based on indices
                    if (cells.size() > Math.max(currentYearIndex, nextYearIndex)) {
                        // Current year estimate (adjusted to match the cell index)
                        String currentYearEps = cells.get(currentYearIndex).text().trim();
                        double currentYearValue = parseDouble(currentYearEps);
                        result.put("eps0", currentYearValue);
                        
                        // Next year estimate (adjusted to match the cell index)
                        String nextYearEps = cells.get(nextYearIndex).text().trim();
                        double nextYearValue = parseDouble(nextYearEps);
                        result.put("eps1", nextYearValue);
                        
                        // Calculate growth rate for year 3 (assuming similar growth)
                        if (currentYearValue > 0 && nextYearValue > 0) {
                            double growthRate = (nextYearValue - currentYearValue) / currentYearValue;
                            double year3Value = nextYearValue * (1 + growthRate);
                            result.put("eps2", Double.parseDouble(String.format("%.2f", year3Value)));
                        } else {
                            result.put("eps2", 0.0);
                        }
                        
                        // Add quarters if available
                        if (cells.size() >= 3) {
                            String currentQtrEps = cells.get(1).text().trim();
                            result.put("currentQtr", parseDouble(currentQtrEps));
                            
                            String nextQtrEps = cells.get(2).text().trim();
                            result.put("nextQtr", parseDouble(nextQtrEps));
                        }
                        
                        // Add the year information
                        result.put("currentYear", currentYear);
                        result.put("nextYear", nextYear);
                        
                        System.out.println("Successfully fetched Yahoo Finance EPS estimates for " + ticker);
                        System.out.println("Current Year EPS: " + currentYearValue);
                        System.out.println("Next Year EPS: " + nextYearValue);
                        
                        return result;
                    }
                }
            }
            
            // If we reach here, we couldn't find the expected data
            System.out.println("Could not find EPS estimates on Yahoo Finance for " + ticker);
            // Try to save the HTML for debugging
            try {
                System.out.println("HTML excerpt: " + doc.select("h3:contains(Earnings Estimate)").parents().first());
            } catch (Exception e) {
                System.out.println("Could not print HTML excerpt");
            }
            
            return new JSONObject();  // Return empty object instead of null
            
        } catch (IOException e) {
            System.err.println("Error fetching Yahoo Finance estimates for " + ticker + ": " + e.getMessage());
            return new JSONObject();  // Return empty object on error
        } catch (Exception e) {
            System.err.println("Unexpected error with Yahoo Finance for " + ticker + ": " + e.getMessage());
            e.printStackTrace();
            return new JSONObject();  // Return empty object on unexpected error
        }
    }
    
    // Helper method to extract year from header text like "Current Year (2025)"
    private static int extractYear(String headerText) {
        try {
            if (headerText.contains("(") && headerText.contains(")")) {
                String yearStr = headerText.substring(headerText.indexOf("(") + 1, headerText.indexOf(")"));
                return Integer.parseInt(yearStr);
            }
        } catch (Exception e) {
            // Silent error - will fall back to calendar year
        }
        return 0;
    }
    
    // Helper method to parse double values from text, handling non-numeric characters
// Helper method to parse double values from text, handling different locale decimal separators
// Helper method to parse double values from text, handling different locale decimal separators
// In the parseDouble method in Estimates.java
private static double parseDouble(String value) {
    try {
        System.out.println("Attempting to parse: '" + value + "'");
        // Check if the value contains a comma
        if (value.contains(",")) {
            System.out.println("Found comma in value: '" + value + "', replacing with period");
            value = value.replace(',', '.');
        }
        
        // Remove any non-numeric characters except decimal point and minus sign
        String cleanValue = value.replaceAll("[^0-9.-]", "");
        
        return Double.parseDouble(cleanValue);
    } catch (NumberFormatException e) {
        System.err.println("Failed to parse value: '" + value + "' - " + e.getMessage());
        e.printStackTrace();
        return 0.0;
    }
}
}