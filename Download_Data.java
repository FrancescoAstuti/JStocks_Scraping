package afin.jstocks;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.FileWriter;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Scanner;

public class Download_Data {
    
    // Reference to the API key
    private static final String API_KEY = API_key.getApiKey();
    
    /**
     * Downloads financial data for all tickers in the watchlist
     * 
     * @param watchlistTable The JTable containing watchlist data
     * @param tableModel The table model for accessing data
     */
    public static void downloadTickerData(JTable watchlistTable, DefaultTableModel tableModel) {
        // Check if there are tickers in the watchlist
        if (tableModel.getRowCount() == 0) {
            JOptionPane.showMessageDialog(null,
                    "No stocks in watchlist to download data for.",
                    "Empty Watchlist",
                    JOptionPane.WARNING_MESSAGE);
            return;
        }

        // Create company_data directory if it doesn't exist
        File dataDir = new File("company_data");
        if (!dataDir.exists()) {
            dataDir.mkdirs();
        }

        // Create progress bar panel - now with 5 API calls per ticker
        JPanel progressPanel = new JPanel(new BorderLayout());
        JProgressBar progressBar = new JProgressBar(0, tableModel.getRowCount() * 5); 
        JLabel statusLabel = new JLabel("Downloading data...");
        progressPanel.add(statusLabel, BorderLayout.NORTH);
        progressPanel.add(progressBar, BorderLayout.CENTER);
        progressPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // Add progress panel to the main window
        Container contentPane = watchlistTable.getRootPane().getContentPane();
        contentPane.add(progressPanel, BorderLayout.SOUTH);
        contentPane.revalidate();
        contentPane.repaint();

        SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
            @Override
            protected Void doInBackground() {
                int rowCount = tableModel.getRowCount();
                int progressCounter = 0;
                
                for (int i = 0; i < rowCount; i++) {
                    int modelRow = watchlistTable.convertRowIndexToModel(i);
                    String ticker = (String) tableModel.getValueAt(modelRow, 1);
                    
                    try {
                        // Create a final copy of the loop counter for use in lambdas
                        final int currentIndex = i + 1;
                        final String currentTicker = ticker;
                        
                        // 1. Update progress for key metrics download
                        final int keyMetricsProgress = progressCounter;
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(keyMetricsProgress);
                            statusLabel.setText("Downloading key metrics for: " + currentTicker + " (" + currentIndex + "/" + rowCount + ")");
                        });

                        // Fetch key metrics data
                        String keyMetricsUrl = String.format("https://financialmodelingprep.com/api/v3/key-metrics/%s?period=annual&apikey=%s", 
                                                        ticker, API_KEY);
                        downloadAndSaveData(keyMetricsUrl, new File(dataDir, ticker + "_Key_Metrics_FY.json"), ticker, "key metrics");
                        
                        // Update progress counter
                        progressCounter++;
                        
                        // 2. Update progress for financial ratios download
                        final int ratiosProgress = progressCounter;
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(ratiosProgress);
                            statusLabel.setText("Downloading financial ratios for: " + currentTicker + " (" + currentIndex + "/" + rowCount + ")");
                        });

                        // Fetch financial ratios data
                        String ratiosUrl = String.format("https://financialmodelingprep.com/api/v3/ratios/%s?period=annual&apikey=%s", 
                                                        ticker, API_KEY);
                        downloadAndSaveData(ratiosUrl, new File(dataDir, ticker + "_Financial_Ratios_FY.json"), ticker, "financial ratios");
                        
                        // Update progress counter
                        progressCounter++;
                        
                        // 3. Update progress for income statement download
                        final int incomeStatementProgress = progressCounter;
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(incomeStatementProgress);
                            statusLabel.setText("Downloading income statement for: " + currentTicker + " (" + currentIndex + "/" + rowCount + ")");
                        });
                        
                        // Fetch income statement data
                        String incomeStatementUrl = String.format("https://financialmodelingprep.com/api/v3/income-statement/%s?period=annual&apikey=%s", 
                                                        ticker, API_KEY);
                        downloadAndSaveData(incomeStatementUrl, new File(dataDir, ticker + "_Income_Statement_FY.json"), ticker, "income statement");
                        
                        // Update progress counter
                        progressCounter++;
                        
                        // 4. Update progress for balance sheet statement download
                        final int balanceSheetProgress = progressCounter;
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(balanceSheetProgress);
                            statusLabel.setText("Downloading balance sheet for: " + currentTicker + " (" + currentIndex + "/" + rowCount + ")");
                        });
                        
                        // Fetch balance sheet statement data
                        String balanceSheetUrl = String.format("https://financialmodelingprep.com/api/v3/balance-sheet-statement/%s?period=annual&apikey=%s", 
                                                        ticker, API_KEY);
                        downloadAndSaveData(balanceSheetUrl, new File(dataDir, ticker + "_Balance_Sheet_Statement_FY.json"), ticker, "balance sheet statement");
                        
                        // Update progress counter
                        progressCounter++;
                        
                        // 5. Update progress for cash flow statement download
                        final int cashFlowProgress = progressCounter;
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(cashFlowProgress);
                            statusLabel.setText("Downloading cash flow statement for: " + currentTicker + " (" + currentIndex + "/" + rowCount + ")");
                        });
                        
                        // Fetch cash flow statement data
                        String cashFlowUrl = String.format("https://financialmodelingprep.com/api/v3/cash-flow-statement/%s?period=annual&apikey=%s", 
                                                        ticker, API_KEY);
                        downloadAndSaveData(cashFlowUrl, new File(dataDir, ticker + "_Cash_Flow_Statement_FY.json"), ticker, "cash flow statement");
                        
                        // Update progress counter
                        progressCounter++;
                        
                        // Small delay to prevent API rate limiting
                        Thread.sleep(500);
                        
                    } catch (Exception e) {
                        e.printStackTrace();
                        System.err.println("Error downloading data for " + ticker + ": " + e.getMessage());
                    }
                }
                return null;
            }

            @Override
            protected void done() {
                // Remove progress panel when done
                contentPane.remove(progressPanel);
                contentPane.revalidate();
                contentPane.repaint();

                JOptionPane.showMessageDialog(
                        watchlistTable,
                        "Data Downloaded",
                        "Download Complete",
                        JOptionPane.INFORMATION_MESSAGE);
                
                System.out.println("All data downloaded to company_data directory");
            }
        };

        worker.execute();
    }
    
    /**
     * Helper method to download and save data from a given URL
     * 
     * @param urlString The URL to download data from
     * @param outputFile The file to save the data to
     * @param ticker The ticker symbol for logging
     * @param dataType The type of data being downloaded for logging
     * @throws Exception If an error occurs during download or save
     */
    private static void downloadAndSaveData(String urlString, File outputFile, String ticker, String dataType) throws Exception {
        URL url = new URL(urlString);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");
        
        int responseCode = connection.getResponseCode();
        if (responseCode == 200) {
            // Read the response
            Scanner scanner = new Scanner(url.openStream());
            String response = scanner.useDelimiter("\\Z").next();
            scanner.close();
            
            // Save to file
            try (FileWriter writer = new FileWriter(outputFile)) {
                writer.write(response);
            }
            
            System.out.println("Downloaded " + dataType + " for " + ticker);
        } else {
            System.err.println("Failed to download " + dataType + " for " + ticker + ": HTTP " + responseCode);
            throw new Exception("HTTP Error: " + responseCode);
        }
    }
}