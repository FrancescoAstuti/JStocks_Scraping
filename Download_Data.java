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
     * Downloads key metrics data for all tickers in the watchlist
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

        // Create progress bar panel
        JPanel progressPanel = new JPanel(new BorderLayout());
        JProgressBar progressBar = new JProgressBar(0, tableModel.getRowCount());
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
                for (int i = 0; i < rowCount; i++) {
                    int modelRow = watchlistTable.convertRowIndexToModel(i);
                    String ticker = (String) tableModel.getValueAt(modelRow, 1);
                    
                    try {
                        // Update progress
                        final int progress = i;
                        final String currentTicker = ticker;
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(progress);
                            statusLabel.setText("Downloading data for: " + currentTicker + " (" + progress + "/" + rowCount + ")");
                        });

                        // Fetch data from Financial Modeling Prep API
                        String urlString = String.format("https://financialmodelingprep.com/api/v3/key-metrics/%s?period=annual&apikey=%s", 
                                                        ticker, API_KEY);
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
                            File outputFile = new File(dataDir, ticker + "_Key_Metrics.json");
                            try (FileWriter writer = new FileWriter(outputFile)) {
                                writer.write(response);
                            }
                            
                            System.out.println("Downloaded data for " + ticker);
                        } else {
                            System.err.println("Failed to download data for " + ticker + ": HTTP " + responseCode);
                        }
                        
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
}