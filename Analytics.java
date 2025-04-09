package afin.jstocks;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.HashMap;
import java.util.Map;
import org.json.JSONArray;
import org.json.JSONObject;
import java.io.FileReader;
import javax.swing.*;
import java.awt.*;
import java.io.File;
import javax.swing.table.DefaultTableModel;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.table.TableRowSorter;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.JScrollPane;

public class Analytics {

    // Modified endpoint to ensure correct API call format
    private static final String ENDPOINT_BASE = "https://financialmodelingprep.com/api/v3/historical-price-full/";
    private static final String API_KEY = API_key.getApiKey();

    /**
     * Analyze price oscillations for all stocks in the watchlist
     */
    public static void analyzeWatchlist() {
        SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
            @Override
            protected Void doInBackground() throws Exception {
                try {
                    System.out.println("Starting analytics...");

                    // Debug API key
                    String maskedKey = API_KEY;
                    if (maskedKey != null && maskedKey.length() > 4) {
                        maskedKey = maskedKey.substring(0, 4) + "...";
                    }
                    System.out.println("Using API key (first 4 chars): " + maskedKey);

                    // Get tickers from watchlist
                    List<String> watchlistTickers = getWatchlistTickers();

                    if (watchlistTickers.isEmpty()) {
                        throw new Exception("No tickers found in watchlist");
                    }

                    System.out.println("Found " + watchlistTickers.size() + " tickers in watchlist");

                    // Get date range (1 year ago from today to today)
                    LocalDate endDate = LocalDate.now();
                    LocalDate startDate = endDate.minusYears(1);
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
                    String formattedEndDate = endDate.format(formatter);
                    String formattedStartDate = startDate.format(formatter);

                    // Results storage
                    Map<String, Integer> upDownCycles = new HashMap<>();     // 10% up then 10% down
                    Map<String, Integer> downUpCycles = new HashMap<>();     // 10% down then 10% up
                    Map<String, Integer> totalCycles = new HashMap<>();      // Total cycles
                    List<String> failedTickers = new ArrayList<>();
                    boolean hitApiLimit = false;

                    // Create progress dialog
                    JDialog progressDialog = createProgressDialog(watchlistTickers.size());
                    JProgressBar progressBar = (JProgressBar) ((JPanel) progressDialog.getContentPane().getComponent(1)).getComponent(0);
                    JLabel statusLabel = (JLabel) ((JPanel) progressDialog.getContentPane().getComponent(0)).getComponent(0);

                    // Process each ticker
                    int count = 0;
                    for (String ticker : watchlistTickers) {
                        final int currentCount = ++count;
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(currentCount);
                            statusLabel.setText("Processing " + ticker + " (" + currentCount + "/" + watchlistTickers.size() + ")");
                        });

                        try {
                            // Add API key to endpoint URL - using the better documented endpoint
                            String endpoint = ENDPOINT_BASE + ticker + "?apikey=" + API_KEY;
                            System.out.println("Fetching data for " + ticker + " from endpoint: "
                                    + endpoint.substring(0, endpoint.indexOf("?") + 8) + "...");

                            JSONObject response = fetchHistoricalDataObj(endpoint);

                            if (response != null && response.has("historical")) {
                                JSONArray priceData = response.getJSONArray("historical");

                                if (priceData != null && priceData.length() > 0) {
                                    // Extract and process price data - reverse to go from oldest to newest
                                    List<Double> closePrices = extractClosePricesFromHistorical(priceData);

                                    if (!closePrices.isEmpty()) {
                                        // Calculate price cycles
                                        System.out.println("Got " + closePrices.size() + " price points for " + ticker);
                                        int[] cycles = calculatePriceCycles(closePrices);

                                        // Store results
                                        upDownCycles.put(ticker, cycles[0]);
                                        downUpCycles.put(ticker, cycles[1]);
                                        totalCycles.put(ticker, cycles[0] + cycles[1]);
                                        System.out.println(ticker + ": UpDown Cycles: " + cycles[0]
                                                + ", DownUp Cycles: " + cycles[1]
                                                + ", Total: " + (cycles[0] + cycles[1]));
                                    } else {
                                        System.out.println("No close prices found for " + ticker);
                                        failedTickers.add(ticker + " (no price data)");
                                    }
                                } else {
                                    System.out.println("Empty historical data array for " + ticker);
                                    failedTickers.add(ticker + " (empty historical data)");
                                }
                            } else {
                                System.out.println("No historical data field in response for " + ticker);
                                failedTickers.add(ticker + " (no historical field in response)");
                            }
                        } catch (Exception e) {
                            String message = e.getMessage();
                            System.out.println("Error processing " + ticker + ": " + message);

                            if (message != null && message.contains("HTTP Error Code: 402")) {
                                hitApiLimit = true;
                                // If we hit the API limit, we'll continue with what we have so far
                                System.out.println("Hit API limit at ticker: " + ticker);
                            } else if (message != null && message.contains("HTTP Error Code: 403")) {
                                throw new Exception("API key unauthorized. Please check your API key.");
                            } else if (message != null && message.contains("HTTP Error Code: 401")) {
                                throw new Exception("API key unauthorized. Please check your API key.");
                            } else {
                                failedTickers.add(ticker + " (" + message + ")");
                            }
                        }

                        // If we've hit API limit, stop processing more tickers
                        if (hitApiLimit) {
                            break;
                        }
                    }

                    // Close progress dialog
                    SwingUtilities.invokeLater(() -> progressDialog.dispose());

                    if (totalCycles.isEmpty()) {
                        SwingUtilities.invokeLater(() -> {
                            JOptionPane.showMessageDialog(null,
                                    "No valid data was retrieved. Please check your API key and try again.\n"
                                    + "Make sure your Financial Modeling Prep API key is valid and has sufficient privileges.",
                                    "Error",
                                    JOptionPane.ERROR_MESSAGE);
                        });
                    } else {
                        // Display results
                        displayResultsInDialog(upDownCycles, downUpCycles, totalCycles,
                                formattedStartDate, formattedEndDate,
                                failedTickers, hitApiLimit);
                    }

                } catch (Exception e) {
                    e.printStackTrace();
                    SwingUtilities.invokeLater(() -> {
                        JOptionPane.showMessageDialog(null,
                                "Error analyzing watchlist: " + e.getMessage(),
                                "Error",
                                JOptionPane.ERROR_MESSAGE);
                    });
                }

                return null;
            }
        };

        worker.execute();
    }

    public static void createDataAnalysisSheet(JTabbedPane tabbedPane) {
        String[] columnNames = {"Column 1", "Column 2", "Column 3"};  // Define your columns
        DefaultTableModel dataAnalysisModel = new DefaultTableModel(columnNames, 0);
        JTable dataAnalysisTable = new JTable(dataAnalysisModel);
        JScrollPane dataAnalysisScrollPane = new JScrollPane(dataAnalysisTable);
        tabbedPane.addTab("Data Analysis", dataAnalysisScrollPane);
    }

    private static JDialog createProgressDialog(int tickerCount) {
        JDialog dialog = new JDialog((Frame) null, "Analyzing Price Cycles", true);
        dialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
        dialog.setLayout(new BorderLayout());

        JPanel statusPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JLabel statusLabel = new JLabel("Initializing...");
        statusPanel.add(statusLabel);

        JPanel progressPanel = new JPanel(new BorderLayout());
        JProgressBar progressBar = new JProgressBar(0, tickerCount);
        progressBar.setStringPainted(true);
        progressPanel.add(progressBar, BorderLayout.CENTER);

        dialog.add(statusPanel, BorderLayout.NORTH);
        dialog.add(progressPanel, BorderLayout.CENTER);
        dialog.setSize(350, 100);
        dialog.setLocationRelativeTo(null);

        // Show dialog in a separate thread to avoid blocking
        SwingUtilities.invokeLater(() -> dialog.setVisible(true));

        return dialog;
    }

    // Updated to handle JSONObject response format
    private static JSONObject fetchHistoricalDataObj(String endpoint) throws Exception {
        URL url = new URL(endpoint);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setConnectTimeout(10000);
        conn.setReadTimeout(10000);

        // Check if the request was successful
        int responseCode = conn.getResponseCode();
        if (responseCode != 200) {
            throw new Exception("Failed to fetch data. HTTP Error Code: " + responseCode);
        }

        // Read the response
        BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
        String inputLine;
        StringBuilder responseContent = new StringBuilder();

        while ((inputLine = in.readLine()) != null) {
            responseContent.append(inputLine);
        }
        in.close();
        conn.disconnect();

        // Parse JSON Object from the response
        return new JSONObject(responseContent.toString());
    }

    // Extract close prices from the historical data array
    private static List<Double> extractClosePricesFromHistorical(JSONArray priceData) {
        List<Double> closePrices = new ArrayList<>();

        // FMP API returns data in reverse chronological order (newest first)
        // We'll reverse it to go from oldest to newest for easier cycle detection
        for (int i = priceData.length() - 1; i >= 0; i--) {
            JSONObject dailyData = priceData.getJSONObject(i);
            if (dailyData.has("close")) {
                closePrices.add(dailyData.getDouble("close"));
            }
        }
        return closePrices;
    }

    /**
     * Calculate price cycles based on 10% up and down movements Returns:
     * [upDownCycles, downUpCycles] - upDownCycles: Number of cycles where price
     * went up 10% and then down 10% - downUpCycles: Number of cycles where
     * price went down 10% and then up 10%
     */
    private static int[] calculatePriceCycles(List<Double> prices) {
        int upDownCycles = 0;  // Up then Down
        int downUpCycles = 0;  // Down then Up

        // Check if we have enough data
        if (prices == null || prices.size() < 2) {
            return new int[]{0, 0};
        }

        enum State {
            NEUTRAL, UP_10, DOWN_10
        }
        State currentState = State.NEUTRAL;

        double referencePrice = prices.get(0);
        double highestPrice = referencePrice;
        double lowestPrice = referencePrice;

        for (int i = 1; i < prices.size(); i++) {
            double currentPrice = prices.get(i);

            switch (currentState) {
                case NEUTRAL:
                    // From neutral state, looking for first 10% move in either direction
                    double percentChangeFromRef = ((currentPrice - referencePrice) / referencePrice) * 100.0;

                    if (percentChangeFromRef >= 10.0) {
                        // Found 10% up move, looking for down move next
                        currentState = State.UP_10;
                        highestPrice = currentPrice;
                    } else if (percentChangeFromRef <= -10.0) {
                        // Found 10% down move, looking for up move next
                        currentState = State.DOWN_10;
                        lowestPrice = currentPrice;
                    }
                    break;

                case UP_10:
                    // After 10% up move, track highest price and look for 10% drop from high
                    if (currentPrice > highestPrice) {
                        highestPrice = currentPrice; // Update high water mark
                    }

                    double dropPercent = ((currentPrice - highestPrice) / highestPrice) * 100.0;
                    if (dropPercent <= -10.0) {
                        // Found 10% down move after up move, complete cycle
                        upDownCycles++;
                        currentState = State.NEUTRAL;
                        referencePrice = currentPrice; // Reset reference for new cycle
                        highestPrice = currentPrice;
                        lowestPrice = currentPrice;
                    }
                    break;

                case DOWN_10:
                    // After 10% down move, track lowest price and look for 10% rise from low
                    if (currentPrice < lowestPrice) {
                        lowestPrice = currentPrice; // Update low water mark
                    }

                    double risePercent = ((currentPrice - lowestPrice) / lowestPrice) * 100.0;
                    if (risePercent >= 10.0) {
                        // Found 10% up move after down move, complete cycle
                        downUpCycles++;
                        currentState = State.NEUTRAL;
                        referencePrice = currentPrice; // Reset reference for new cycle
                        highestPrice = currentPrice;
                        lowestPrice = currentPrice;
                    }
                    break;
            }
        }

        return new int[]{upDownCycles, downUpCycles};
    }

    public static double calculateVolatilityScore(String ticker) {
        try {
            String endpoint = ENDPOINT_BASE + ticker + "?apikey=" + API_KEY;
            JSONObject response = fetchHistoricalDataObj(endpoint);

            if (response != null && response.has("historical")) {
                JSONArray priceData = response.getJSONArray("historical");

                if (priceData != null && priceData.length() > 0) {
                    List<Double> closePrices = extractClosePricesFromHistorical(priceData);

                    if (closePrices.size() >= 30) { // Richiede almeno 30 punti di dati
                        // Calcola i rendimenti giornalieri logaritmici
                        List<Double> logReturns = new ArrayList<>();
                        for (int i = 1; i < closePrices.size(); i++) {
                            double current = closePrices.get(i);
                            double previous = closePrices.get(i - 1);
                            if (previous > 0 && current > 0) {
                                // Utilizzo dei rendimenti logaritmici che sono più comuni in finanza
                                logReturns.add(Math.log(current / previous));
                            }
                        }

                        // Calcola la deviazione standard
                        double sum = 0.0;
                        for (Double ret : logReturns) {
                            sum += ret;
                        }
                        double mean = sum / logReturns.size();

                        double squaredDiffSum = 0.0;
                        for (Double ret : logReturns) {
                            squaredDiffSum += Math.pow(ret - mean, 2);
                        }
                        double stdDev = Math.sqrt(squaredDiffSum / logReturns.size());

                        // Annualizza (assumendo dati giornalieri, tipicamente si usa √252)
                        double annualizedStdDev = stdDev * Math.sqrt(252);

                        // Converti in un punteggio da 0-10
                        // Nella finanza, una volatilità annualizzata:
                        // - < 15% è considerata bassa (1-3)
                        // - 15%-30% è considerata media (4-6)
                        // - > 30% è considerata alta (7-10)
                        // Formula: mappa le volatilità del 15%, 30% e 60% rispettivamente a 3, 6 e 10
                        double volatilityScore = (annualizedStdDev * 100) / 6;

                        // Limita il punteggio tra 0 e 10 e arrotonda a 1 decimale
                        return Math.min(10.0, Math.round(volatilityScore * 10) / 10.0);
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("Errore nel calcolo della volatilità per " + ticker + ": " + e.getMessage());
        }

        return 0.0; // Punteggio predefinito se il calcolo fallisce
    }

    private static void displayResultsInDialog(Map<String, Integer> upDownCycles,
            Map<String, Integer> downUpCycles,
            Map<String, Integer> totalCycles,
            String startDate,
            String endDate,
            List<String> failedTickers,
            boolean hitApiLimit) {
        SwingUtilities.invokeLater(() -> {
            JDialog resultsDialog = new JDialog((Frame) null, "Price Cycle Analysis", true);
            resultsDialog.setLayout(new BorderLayout());

            // Create tabbed pane for results and errors
            JTabbedPane tabbedPane = new JTabbedPane();

            // Tab 1: Results
            JPanel resultsPanel = new JPanel(new BorderLayout());

            JPanel headerPanel = new JPanel(new BorderLayout());
            headerPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
            JLabel titleLabel = new JLabel("Price Cycle Analysis: " + startDate + " to " + endDate);
            titleLabel.setFont(new java.awt.Font("Dialog", java.awt.Font.BOLD, 14));
            headerPanel.add(titleLabel, BorderLayout.NORTH);

            if (hitApiLimit) {
                JLabel warningLabel = new JLabel("<html><b>Note:</b> API request limit reached. Some tickers could not be processed.</html>");
                warningLabel.setForeground(new java.awt.Color(200, 0, 0));
                warningLabel.setBorder(BorderFactory.createEmptyBorder(5, 0, 0, 0));
                headerPanel.add(warningLabel, BorderLayout.CENTER);
            }

            resultsPanel.add(headerPanel, BorderLayout.NORTH);

            // Create table model
            String[] columns = {"Ticker", "UpDown Cycles", "DownUp Cycles", "Total Cycles", "Volatility Rating"};
            DefaultTableModel model = new DefaultTableModel(columns, 0) {
                @Override
                public Class<?> getColumnClass(int column) {
                    if (column >= 1 && column <= 4) {
                        return Double.class;
                    }
                    return String.class;
                }

                @Override
                public boolean isCellEditable(int row, int column) {
                    return false;
                }
            };

            // Add data
            for (String ticker : totalCycles.keySet()) {
                int upDown = upDownCycles.getOrDefault(ticker, 0);
                int downUp = downUpCycles.getOrDefault(ticker, 0);
                int total = totalCycles.get(ticker);

                // Calculate volatility rating (simple scale 1-10)
                // Assuming 5+ cycles in a year is extremely volatile (10/10)
                int volatilityRating = Math.min(10, (total * 2));

                model.addRow(new Object[]{
                    ticker,
                    upDown,
                    downUp,
                    total,
                    volatilityRating
                });
            }

            // Create table
            JTable resultsTable = new JTable(model);

            // Add sorting capability
            TableRowSorter<DefaultTableModel> sorter = new TableRowSorter<>(model);
            resultsTable.setRowSorter(sorter);

            JScrollPane scrollPane = new JScrollPane(resultsTable);
            resultsPanel.add(scrollPane, BorderLayout.CENTER);

            // Add note label
            JPanel notePanel = new JPanel(new BorderLayout());
            notePanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
            JLabel noteLabel = new JLabel("<html><b>Note on Price Cycles:</b><br>"
                    + "- UpDown Cycle: Price rises by 10% then falls by 10%<br>"
                    + "- DownUp Cycle: Price falls by 10% then rises by 10%<br>"
                    + "- Total Cycles: Sum of both cycle types<br>"
                    + "- Volatility Rating: Scale 1-10 based on cycle frequency</html>");
            notePanel.add(noteLabel, BorderLayout.CENTER);
            resultsPanel.add(notePanel, BorderLayout.SOUTH);

            // Add the results panel to the tabbed pane
            tabbedPane.addTab("Results", resultsPanel);

            // Tab 2: Failed tickers
            if (!failedTickers.isEmpty()) {
                JPanel errorPanel = new JPanel(new BorderLayout());
                JTextArea errorText = new JTextArea();
                errorText.setEditable(false);
                StringBuilder sb = new StringBuilder();
                sb.append("The following tickers could not be processed:\n\n");
                for (String ticker : failedTickers) {
                    sb.append(ticker).append("\n");
                }
                errorText.setText(sb.toString());
                JScrollPane errorScroll = new JScrollPane(errorText);
                errorPanel.add(errorScroll, BorderLayout.CENTER);
                tabbedPane.addTab("Failed Tickers (" + failedTickers.size() + ")", errorPanel);
            }

            // Add tabbed pane to dialog
            resultsDialog.add(tabbedPane, BorderLayout.CENTER);

            // Add export button
            JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            JButton exportButton = new JButton("Export to Excel");
            exportButton.addActionListener(e -> {
                try {
                    exportResultsToExcel(upDownCycles, downUpCycles, totalCycles,
                            startDate, endDate, failedTickers, hitApiLimit);
                    JOptionPane.showMessageDialog(resultsDialog,
                            "Results exported successfully!",
                            "Export Complete",
                            JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(resultsDialog,
                            "Error exporting results: " + ex.getMessage(),
                            "Export Error",
                            JOptionPane.ERROR_MESSAGE);
                }
            });
            buttonPanel.add(exportButton);
            resultsDialog.add(buttonPanel, BorderLayout.SOUTH);

            // Show dialog
            resultsDialog.setSize(600, 500);
            resultsDialog.setLocationRelativeTo(null);
            resultsDialog.setVisible(true);
        });
    }

    private static void exportResultsToExcel(Map<String, Integer> upDownCycles,
            Map<String, Integer> downUpCycles,
            Map<String, Integer> totalCycles,
            String startDate,
            String endDate,
            List<String> failedTickers,
            boolean hitApiLimit) throws IOException {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save Analysis Results");
        fileChooser.setSelectedFile(new File("price_cycle_analysis.xlsx"));

        if (fileChooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();
            if (!file.getName().endsWith(".xlsx")) {
                file = new File(file.getAbsolutePath() + ".xlsx");
            }

            try (Workbook workbook = new XSSFWorkbook()) {
                // Results sheet
                Sheet sheet = workbook.createSheet("Price Cycle Analysis");

                // Create styles
                CellStyle headerStyle = workbook.createCellStyle();
                headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                org.apache.poi.ss.usermodel.Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                headerStyle.setFont(headerFont);

                CellStyle warningStyle = workbook.createCellStyle();
                org.apache.poi.ss.usermodel.Font warningFont = workbook.createFont();
                warningFont.setColor(IndexedColors.RED.getIndex());
                warningFont.setBold(true);
                warningStyle.setFont(warningFont);

                // Add title and date info
                Row titleRow = sheet.createRow(0);
                Cell titleCell = titleRow.createCell(0);
                titleCell.setCellValue("Price Cycle Analysis");
                titleCell.setCellStyle(headerStyle);

                Row dateRow = sheet.createRow(1);
                dateRow.createCell(0).setCellValue("Period: " + startDate + " to " + endDate);

                // Add API limit warning if needed
                if (hitApiLimit) {
                    Row warningRow = sheet.createRow(2);
                    Cell warningCell = warningRow.createCell(0);
                    warningCell.setCellValue("NOTE: API request limit reached. Some tickers could not be processed.");
                    warningCell.setCellStyle(warningStyle);
                }

                // Add headers
                Row headerRow = sheet.createRow(hitApiLimit ? 3 : 2);
                String[] headers = {"Ticker", "UpDown Cycles", "DownUp Cycles", "Total Cycles", "Volatility Rating"};
                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers[i]);
                    cell.setCellStyle(headerStyle);
                }

                // Add data
                int rowNum = headerRow.getRowNum() + 1;
                for (String ticker : totalCycles.keySet()) {
                    int upDown = upDownCycles.getOrDefault(ticker, 0);
                    int downUp = downUpCycles.getOrDefault(ticker, 0);
                    int total = totalCycles.get(ticker);

                    // Calculate volatility rating (simple scale 1-10)
                    // Assuming 5+ cycles in a year is extremely volatile (10/10)
                    int volatilityRating = Math.min(10, (total * 2));

                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(ticker);
                    row.createCell(1).setCellValue(upDown);
                    row.createCell(2).setCellValue(downUp);
                    row.createCell(3).setCellValue(total);
                    row.createCell(4).setCellValue(volatilityRating);
                }

                // Add notes
                Row noteRow = sheet.createRow(rowNum + 1);
                noteRow.createCell(0).setCellValue("Note on Price Cycles:");

                Row note1Row = sheet.createRow(rowNum + 2);
                note1Row.createCell(0).setCellValue("- UpDown Cycle: Price rises by 10% then falls by 10%");

                Row note2Row = sheet.createRow(rowNum + 3);
                note2Row.createCell(0).setCellValue("- DownUp Cycle: Price falls by 10% then rises by 10%");

                Row note3Row = sheet.createRow(rowNum + 4);
                note3Row.createCell(0).setCellValue("- Volatility Rating: Scale of 1-10, where 10 is highly volatile (5+ cycles per year)");

                // Auto size columns
                for (int i = 0; i < headers.length; i++) {
                    sheet.autoSizeColumn(i);
                }

                // Create Failed Tickers sheet if needed
                if (!failedTickers.isEmpty()) {
                    Sheet failedSheet = workbook.createSheet("Failed Tickers");

                    // Add header
                    Row fHeaderRow = failedSheet.createRow(0);
                    Cell fHeaderCell = fHeaderRow.createCell(0);
                    fHeaderCell.setCellValue("Failed Tickers");
                    fHeaderCell.setCellStyle(headerStyle);

                    // Add explanation row
                    Row explainRow = failedSheet.createRow(1);
                    explainRow.createCell(0).setCellValue("The following tickers could not be processed:");

                    // Add tickers
                    for (int i = 0; i < failedTickers.size(); i++) {
                        Row tickerRow = failedSheet.createRow(i + 3);
                        tickerRow.createCell(0).setCellValue(failedTickers.get(i));
                    }

                    failedSheet.autoSizeColumn(0);
                }

                // Write to file
                try (FileOutputStream fileOut = new FileOutputStream(file)) {
                    workbook.write(fileOut);
                }
            }
        }
    }

    private static List<String> getWatchlistTickers() throws Exception {
        List<String> tickers = new ArrayList<>();

        try {
            BufferedReader reader = new BufferedReader(new FileReader("watchlist.json"));
            StringBuilder content = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                content.append(line);
            }
            reader.close();

            JSONArray watchlist = new JSONArray(content.toString());
            System.out.println("Watchlist JSON has " + watchlist.length() + " entries");

            for (int i = 0; i < watchlist.length(); i++) {
                JSONObject stock = watchlist.getJSONObject(i);
                if (stock.has("ticker")) {
                    tickers.add(stock.getString("ticker"));
                }
            }

            if (tickers.isEmpty()) {
                throw new Exception("No tickers found in watchlist.json");
            }

        } catch (Exception e) {
            throw new Exception("Error reading watchlist: " + e.getMessage());
        }

        return tickers;
    }

    /**
     * Calcola il punteggio di volatilità per un singolo stock utilizzando la
     * deviazione standard annualizzata dei rendimenti giornalieri
     *
     * @param ticker Simbolo del ticker dello stock
     * @return Punteggio di volatilità su una scala da 0-10
     */
}
