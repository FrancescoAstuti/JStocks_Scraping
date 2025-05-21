package afin.jstocks;

import afin.jstocks.CellsFormat;

import javax.swing.*;
import javax.swing.event.RowSorterEvent;
import javax.swing.event.RowSorterListener;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;
import javax.swing.table.TableRowSorter;
import java.awt.*;
// import java.awt.Font; // No longer needed as individual import if qualified
import java.awt.event.ActionListener; // Keep this specific import
import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.HttpURLConnection;
import java.net.URL;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;
import org.json.JSONArray;
import org.json.JSONObject;
import java.util.Calendar;
import javax.swing.JProgressBar;
import javax.swing.BorderFactory;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import javax.swing.UIManager;
import javax.swing.UIManager.LookAndFeelInfo;
// import java.awt.Color; // Ambiguity source, qualify usages
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
// import org.apache.poi.xssf.usermodel.XSSFColor; // Already imported in CellsFormat if needed there
import java.io.FileOutputStream;

public class Watchlist {

    private JTable watchlistTable;
    private DefaultTableModel tableModel;
    private static final String API_KEY = API_key.getApiKey();
    private static final String COLUMN_SETTINGS_FILE = "column_settings.properties";
    private static final String REFRESH_STATE_FILE = "refresh_state.properties"; // For saving refresh index
    private JPanel columnControlPanel;

    // Instance variables for resumable refresh
    private int lastSuccessfullyProcessedIndex = -1;
    private SwingWorker<Void, Void> refreshDataWorker;
    private JButton refreshButton; 
    private JButton resumeButton;  
    private JButton cancelButton;  

    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");

    // Promoted to instance variables
    private String[] dynamicColumnNames;
    private String[] peForwardColumnNames;


    private String[] getDynamicColumnNames() {
        int currentYear = Calendar.getInstance().get(Calendar.YEAR);
        String[] columnNames = new String[3];
        for (int i = 0; i < 3; i++) {
            columnNames[i] = "EPS" + (currentYear + i);
        }
        return columnNames;
    }
    
    private String[] getPEForwardColumnNames() {
        int currentYear = Calendar.getInstance().get(Calendar.YEAR);
        String[] columnNames = new String[3];
        for (int i = 0; i < 3; i++) {
            columnNames[i] = "PE FWD" + (currentYear + i);
        }
        return columnNames;
    }


    // Helper method to find column index by name
    private int getColumnIndexByName(String columnName) {
        for (int i = 0; i < tableModel.getColumnCount(); i++) {
            if (tableModel.getColumnName(i).equalsIgnoreCase(columnName)) {
                return i;
            }
        }
        // System.err.println("Warning: Column '" + columnName + "' not found in table model for getColumnIndexByName.");
        return -1; 
    }


    public void createAndShowGUI() {

        try {
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    UIManager.setLookAndFeel(info.getClassName());
                    UIManager.put("control", new java.awt.Color(240, 240, 240));
                    UIManager.put("info", new java.awt.Color(242, 242, 189));
                    UIManager.put("nimbusBase", new java.awt.Color(51, 98, 140));
                    UIManager.put("nimbusBlueGrey", new java.awt.Color(169, 176, 190));
                    UIManager.put("nimbusSelectionBackground", new java.awt.Color(104, 93, 156));
                    break;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }

        JFrame frame = new JFrame("Watchlist");
        frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
        frame.setSize(1024, 768); 

        JPanel mainPanel = new JPanel(new BorderLayout());

        this.dynamicColumnNames = getDynamicColumnNames();
        this.peForwardColumnNames = getPEForwardColumnNames();
        
        tableModel = new DefaultTableModel(new Object[]{
            "Name", "Ticker", "Refresh Date", "Price", "PE TTM", "PB TTM", "Div. yield %", 
            "Payout Ratio", "Graham Number", "PB Avg", "PE Avg",
            "EPS TTM", "ROE TTM", "A-Score",
            this.dynamicColumnNames[0], this.dynamicColumnNames[1], this.dynamicColumnNames[2], 
            "Debt to Equity", "EPS Growth 1", "Current Ratio", "Quick Ratio",
            "EPS Growth 2", "EPS Growth 3", "DE Avg", "Industry", "ROE Avg", "P/FCF", "PFCF Avg",
            this.peForwardColumnNames[0], this.peForwardColumnNames[1], this.peForwardColumnNames[2], "Volatility","PR Avg"}, 0) {

            @Override
            public Class<?> getColumnClass(int columnIndex) {
                String columnName = getColumnName(columnIndex); 
                if ("Refresh Date".equals(columnName) || "Name".equals(columnName) || "Ticker".equals(columnName) || "Industry".equals(columnName)) {
                    return String.class;
                }
                return Double.class;
            }
        };

        watchlistTable = new JTable(tableModel);

        TableRowSorter<DefaultTableModel> sorter = new TableRowSorter<>(tableModel);
        setupTableSorter(sorter);
        watchlistTable.setRowSorter(sorter);

        watchlistTable.getTableHeader().setReorderingAllowed(true);

        CellsFormat.CustomCellRenderer renderer = new CellsFormat.CustomCellRenderer();
        for (int i = 0; i < watchlistTable.getColumnCount(); i++) {
            watchlistTable.getColumnModel().getColumn(i).setCellRenderer(renderer);
        }
        
        JPopupMenu popupMenu = new JPopupMenu();
        JMenuItem refreshStockItem = new JMenuItem("Refresh Stock");
        refreshStockItem.addActionListener(e -> {
            int row = watchlistTable.getSelectedRow();
            if (row != -1) {
                int modelRow = watchlistTable.convertRowIndexToModel(row);
                String ticker = (String) tableModel.getValueAt(modelRow, getColumnIndexByName("Ticker"));
                refreshSingleStock(ticker, modelRow);
            }
        });
        popupMenu.add(refreshStockItem);

        JMenuItem clearStockItem = new JMenuItem("Clear Stock Values");
        clearStockItem.addActionListener(e -> {
            int row = watchlistTable.getSelectedRow();
            if (row != -1) {
                int modelRow = watchlistTable.convertRowIndexToModel(row);
                clearStockValues(modelRow);
            }
        });
        popupMenu.add(clearStockItem);

        watchlistTable.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 1 && e.getButton() == MouseEvent.BUTTON1) {
                    int row = watchlistTable.rowAtPoint(e.getPoint());
                    int col = watchlistTable.columnAtPoint(e.getPoint());
                    if (row >= 0 && col == getColumnIndexByName("Ticker")) { 
                        String ticker = (String) watchlistTable.getValueAt(row, col);
                        String companyName = (String) watchlistTable.getValueAt(row, getColumnIndexByName("Name"));
                        CompanyOverview.showCompanyOverview(ticker, companyName);
                    }
                }
            }

            @Override
            public void mousePressed(MouseEvent e) { showPopupIfNeeded(e); }
            @Override
            public void mouseReleased(MouseEvent e) { showPopupIfNeeded(e); }

            private void showPopupIfNeeded(MouseEvent e) {
                if (e.isPopupTrigger()) {
                    int row = watchlistTable.rowAtPoint(e.getPoint());
                    if (row >= 0) {
                        watchlistTable.setRowSelectionInterval(row, row);
                        popupMenu.show(e.getComponent(), e.getX(), e.getY());
                    }
                }
            }
        });
        
        JScrollPane scrollPane = new JScrollPane(
            watchlistTable,
            JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
            JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED
        );
        watchlistTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        mainPanel.add(scrollPane, BorderLayout.CENTER);

        setupColumnControlPanel();
        mainPanel.add(columnControlPanel, BorderLayout.WEST);

        JPanel buttonPanel = new JPanel();
        setupButtonPanel(buttonPanel);
        mainPanel.add(buttonPanel, BorderLayout.SOUTH);

        frame.add(mainPanel);
        frame.setLocationRelativeTo(null);

        frame.addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosing(java.awt.event.WindowEvent windowEvent) {
                boolean closeConfirmed = false;
                boolean workerWasActiveAndCancelled = false;

                if (refreshDataWorker != null && !refreshDataWorker.isDone()) {
                    int confirmCloseDuringRefresh = JOptionPane.showConfirmDialog(
                        frame,
                        "A refresh operation is in progress. Are you sure you want to close?\nThis will cancel the current refresh.",
                        "Confirm Close During Refresh",
                        JOptionPane.YES_NO_OPTION,
                        JOptionPane.WARNING_MESSAGE
                    );
                    if (confirmCloseDuringRefresh == JOptionPane.YES_OPTION) {
                        refreshDataWorker.cancel(true); 
                        // The worker's done() method will call saveRefreshState().
                        workerWasActiveAndCancelled = true; 
                        closeConfirmed = true; 
                    } else {
                        return; 
                    }
                } else {
                    int confirmStandardClose = JOptionPane.showOptionDialog(
                            frame, "Are you sure you want to close this window?",
                            "Close Confirmation", JOptionPane.YES_NO_OPTION,
                            JOptionPane.QUESTION_MESSAGE, null, null, null);
                    if (confirmStandardClose == JOptionPane.YES_OPTION) {
                        closeConfirmed = true;
                    }
                }

                if (closeConfirmed) {
                    saveColumnSettings();
                    saveWatchlist(); // Save current table data regardless of refresh state
                    if (!workerWasActiveAndCancelled && (refreshDataWorker == null || refreshDataWorker.isDone())) {
                        // If worker was not active or already finished, save current refresh state
                        saveRefreshState();
                    }
                    frame.dispose(); 
                }
            }
        });

        SwingUtilities.invokeLater(() -> {
            loadColumnSettings(); 
            loadWatchlist();
            loadRefreshState(); // Load the refresh state
            loadSortOrder(sorter); 
            // updateRefreshButtonsState(); // Called at the end of loadRefreshState
        });

        frame.setVisible(true);
    }

    private void setupTableSorter(TableRowSorter<DefaultTableModel> sorter) {
        for (int i = 0; i < tableModel.getColumnCount(); i++) {
            if (tableModel.getColumnClass(i) == Double.class) {
                 sorter.setComparator(i, Comparator.comparingDouble(o -> (Double) o));
            }
        }
        sorter.addRowSorterListener(e -> {
            if (e.getType() == RowSorterEvent.Type.SORT_ORDER_CHANGED) {
                saveSortOrder(sorter.getSortKeys());
            }
        });
    }

    private void setupColumnControlPanel() {
        columnControlPanel = new JPanel();
        columnControlPanel.setLayout(new BoxLayout(columnControlPanel, BoxLayout.Y_AXIS));
        TableColumnModel columnModel = watchlistTable.getColumnModel();
        List<JCheckBox> checkBoxes = new ArrayList<>();
        for (int i = 0; i < columnModel.getColumnCount(); i++) {
            String columnName = tableModel.getColumnName(i); 
            JCheckBox checkBox = new JCheckBox(columnName, true);
            checkBox.addActionListener(e -> toggleColumnVisibility(columnName, checkBox.isSelected()));
            checkBoxes.add(checkBox);
        }
        checkBoxes.sort(Comparator.comparing(JCheckBox::getText));
        for (JCheckBox checkBox : checkBoxes) {
            columnControlPanel.add(checkBox);
        }
    }

    private void setupButtonPanel(JPanel buttonPanel) {
        JButton addButton = new JButton("Add Stock");
        JButton deleteButton = new JButton("Delete Stock");
        refreshButton = new JButton("Refresh All"); 
        resumeButton = new JButton("Resume Refresh"); 
        cancelButton = new JButton("Cancel Refresh"); 
        JButton exportButton = new JButton("Export XLSX");
        JButton analyticsButton = new JButton("Analytics");
        JButton dataButton = new JButton("Data");

        addButton.addActionListener(e -> addStock());
        deleteButton.addActionListener(e -> deleteStock());
        refreshButton.addActionListener(e -> startRefreshDataWorker(false)); 
        resumeButton.addActionListener(e -> startRefreshDataWorker(true));   
        cancelButton.addActionListener(e -> {
            if (refreshDataWorker != null && !refreshDataWorker.isDone()) {
                refreshDataWorker.cancel(true);
            }
        });
        exportButton.addActionListener(e -> exportToExcel());
        analyticsButton.addActionListener(e -> Analytics.analyzeWatchlist()); 
        dataButton.addActionListener(e -> Download_Data.downloadTickerData(watchlistTable, tableModel));

        buttonPanel.add(addButton);
        buttonPanel.add(deleteButton);
        buttonPanel.add(refreshButton);
        buttonPanel.add(resumeButton);
        buttonPanel.add(cancelButton);
        buttonPanel.add(exportButton);
        buttonPanel.add(analyticsButton); 
        buttonPanel.add(dataButton);
    }
    
    private void updateRefreshButtonsState() {
        SwingUtilities.invokeLater(() -> { 
            boolean workerActive = refreshDataWorker != null && !refreshDataWorker.isDone();
            refreshButton.setEnabled(!workerActive);
            boolean canResume = lastSuccessfullyProcessedIndex >= 0 && 
                                lastSuccessfullyProcessedIndex < tableModel.getRowCount() - 1;
            resumeButton.setEnabled(!workerActive && canResume);
            cancelButton.setEnabled(workerActive);
        });
    }
    
    private void saveRefreshState() {
        Properties properties = new Properties();
        properties.setProperty("lastSuccessfullyProcessedIndex", String.valueOf(this.lastSuccessfullyProcessedIndex));
        try (FileOutputStream out = new FileOutputStream(REFRESH_STATE_FILE)) {
            properties.store(out, "Watchlist Refresh State");
            System.out.println("Refresh state saved. Last processed index: " + this.lastSuccessfullyProcessedIndex);
        } catch (IOException e) {
            System.err.println("Error saving refresh state: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void loadRefreshState() {
        Properties properties = new Properties();
        File stateFile = new File(REFRESH_STATE_FILE);
        if (stateFile.exists()) {
            try (FileInputStream in = new FileInputStream(stateFile)) {
                properties.load(in);
                this.lastSuccessfullyProcessedIndex = Integer.parseInt(properties.getProperty("lastSuccessfullyProcessedIndex", "-1"));
                System.out.println("Refresh state loaded. Last processed index: " + this.lastSuccessfullyProcessedIndex);
            } catch (IOException | NumberFormatException e) {
                System.err.println("Error loading refresh state or invalid format: " + e.getMessage());
                this.lastSuccessfullyProcessedIndex = -1; // Default to -1 on error
            }
        } else {
            this.lastSuccessfullyProcessedIndex = -1; // Default if file doesn't exist
            System.out.println("Refresh state file not found. Initializing last processed index to -1.");
        }
        updateRefreshButtonsState(); // Update buttons after loading state
    }


    private void startRefreshDataWorker(boolean isResuming) {
        if (refreshDataWorker != null && !refreshDataWorker.isDone()) {
            JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), 
                "A refresh operation is already in progress.", "In Progress", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
         if (tableModel.getRowCount() == 0) {
            JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Watchlist is empty. Nothing to refresh.", "Empty Watchlist", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        JPanel progressPanel = new JPanel(new BorderLayout());
        JProgressBar progressBar = new JProgressBar(0, 100); 
        JLabel statusLabel = new JLabel("Initializing refresh...");
        progressPanel.add(statusLabel, BorderLayout.NORTH);
        progressPanel.add(progressBar, BorderLayout.CENTER);
        progressPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        Container contentPane = watchlistTable.getRootPane().getContentPane();
        contentPane.add(progressPanel, BorderLayout.SOUTH);
        contentPane.revalidate();
        contentPane.repaint();

        this.refreshDataWorker = new SwingWorker<Void, Void>() {
            private int workerInternalLastProcessedIndex = -1; 
            private int totalItemsForThisRunGlobal = 0; 

            @Override
            protected Void doInBackground() {
                updateRefreshButtonsState(); 

                int rowCount = tableModel.getRowCount();
                int startIndex = 0;

                if (isResuming && Watchlist.this.lastSuccessfullyProcessedIndex >= 0 && Watchlist.this.lastSuccessfullyProcessedIndex < rowCount - 1) {
                    startIndex = Watchlist.this.lastSuccessfullyProcessedIndex + 1;
                    workerInternalLastProcessedIndex = Watchlist.this.lastSuccessfullyProcessedIndex;
                    System.out.println("Resuming refresh from index: " + startIndex);
                } else { // Not resuming, or invalid resume state, so start fresh
                    Watchlist.this.lastSuccessfullyProcessedIndex = -1; 
                    saveRefreshState(); // Clear any persisted resumable state
                    workerInternalLastProcessedIndex = -1;
                    System.out.println("Starting full refresh from beginning. Resetting last processed index.");
                }


                totalItemsForThisRunGlobal = rowCount - startIndex;
                
                if (totalItemsForThisRunGlobal <= 0) {
                     SwingUtilities.invokeLater(() -> {
                        if (rowCount > 0) { 
                           statusLabel.setText("Already up-to-date or nothing to resume.");
                           progressBar.setValue(progressBar.getMaximum());
                        } else {
                           statusLabel.setText("Watchlist is empty.");
                        }
                     });
                    return null; 
                }
                
                SwingUtilities.invokeLater(() -> {
                     progressBar.setMaximum(totalItemsForThisRunGlobal);
                     progressBar.setValue(0);
                });

                for (int i = startIndex; i < rowCount; i++) {
                    if (isCancelled()) {
                        Watchlist.this.lastSuccessfullyProcessedIndex = workerInternalLastProcessedIndex; 
                        System.out.println("Refresh cancelled by user. Last successfully processed index: " + Watchlist.this.lastSuccessfullyProcessedIndex);
                        break;
                    }

                    int modelRow = watchlistTable.convertRowIndexToModel(i); 
                    String ticker = (String) tableModel.getValueAt(modelRow, getColumnIndexByName("Ticker"));
                    
                    final int progressValue = i - startIndex + 1;
                    final String currentTickerForLabel = ticker;
                    SwingUtilities.invokeLater(() -> {
                        statusLabel.setText("Refreshing: " + currentTickerForLabel + " (" + progressValue + "/" + totalItemsForThisRunGlobal + ")");
                        progressBar.setValue(progressValue);
                    });

                    try {
                        JSONObject stockData = fetchStockData(ticker);
                        JSONObject ratios = fetchStockRatios(ticker);
                        JSONObject epsEstimates = Estimates.fetchEpsEstimates(ticker);

                        if (stockData != null) { 
                            final double price = convertGbxToGbp(round(stockData.optDouble("price", 0.0), 2), ticker);
                            double peTtm = round(ratios.optDouble("peRatioTTM", 0.0), 2);
                            double pbTtm = round(ratios.optDouble("pbRatioTTM", 0.0), 2);
                            double epsTtm = peTtm != 0 ? round((1 / peTtm) * price, 2) : 0.0;
                            double roeTtm = round(ratios.optDouble("roeTTM", 0.0), 2);
                            double dividendYieldTTM = KeyMetricsTTM.fetchDividendYield(ticker);
                            double payoutRatioVal = KeyMetricsTTM.fetchPayoutRatio(ticker); 
                            double debtToEquityVal = KeyMetricsTTM.fetchDebtToEquity(ticker); 
                            double epsCurrentYear = epsEstimates != null ? round(epsEstimates.optDouble("eps0", 0.0), 2) : 0.0;
                            double epsNextYear = epsEstimates != null ? round(epsEstimates.optDouble("eps1", 0.0), 2) : 0.0;
                            double epsYear3 = epsEstimates != null ? round(epsEstimates.optDouble("eps2", 0.0), 2) : 0.0;
                            double currentRatioVal = KeyMetricsTTM.fetchCurrentRatio(ticker); 
                            double quickRatioVal = 0.0; 

                            double epsGrowth1 = calculateEpsGrowth1(epsCurrentYear, epsTtm);
                            double epsGrowth2 = calculateEpsGrowth2(epsCurrentYear, epsNextYear);
                            double epsGrowth3 = calculateEpsGrowth3(epsYear3, epsNextYear);
                            double pbAvg = fetchAveragePB(ticker);
                            double peForward1 = calculatePEForward(price, epsCurrentYear);
                            double peForward2 = calculatePEForward(price, epsNextYear);
                            double peForward3 = calculatePEForward(price, epsYear3);
                            double peAvg = fetchAveragePE(ticker);
                            double priceToFCF_TTM = TTM_Ratios.getPriceToFreeCashFlowRatioTTM(ticker);
                            double PriceToFCF_Avg = Ratios.fetchPriceToFreeCashFlowAverage(ticker);
                            double roeAvg = fetchAverageROE(ticker);
                            double grahamNumber = calculateGrahamNumber(price, peAvg, pbAvg, epsTtm, pbTtm);
                            double deAvg = Ratios.fetchDebtToEquityAverage(ticker);
                            String industry = CompanyOverview.fetchIndustry(ticker);
                            double prAvg = Ratios.fetchPayoutRatioAverage(ticker);
                            double aScore = calculateAScore(pbAvg, pbTtm, peAvg, peTtm, payoutRatioVal, debtToEquityVal, roeTtm, roeAvg, dividendYieldTTM, deAvg, epsGrowth1, epsGrowth2, epsGrowth3, currentRatioVal, quickRatioVal, grahamNumber, price, priceToFCF_TTM, PriceToFCF_Avg, prAvg);
                            double volatilityScore = Analytics.calculateVolatilityScore(ticker);
                            String refreshDateStr = LocalDate.now().format(DATE_FORMATTER);

                            final int finalModelRow = modelRow; 
                            SwingUtilities.invokeLater(() -> {
                                tableModel.setValueAt(refreshDateStr, finalModelRow, getColumnIndexByName("Refresh Date"));
                                tableModel.setValueAt(price, finalModelRow, getColumnIndexByName("Price"));
                                tableModel.setValueAt(peTtm, finalModelRow, getColumnIndexByName("PE TTM"));
                                tableModel.setValueAt(pbTtm, finalModelRow, getColumnIndexByName("PB TTM"));
                                tableModel.setValueAt(dividendYieldTTM, finalModelRow, getColumnIndexByName("Div. yield %"));
                                tableModel.setValueAt(payoutRatioVal, finalModelRow, getColumnIndexByName("Payout Ratio"));
                                tableModel.setValueAt(grahamNumber, finalModelRow, getColumnIndexByName("Graham Number"));
                                tableModel.setValueAt(pbAvg, finalModelRow, getColumnIndexByName("PB Avg"));
                                tableModel.setValueAt(peAvg, finalModelRow, getColumnIndexByName("PE Avg"));
                                tableModel.setValueAt(epsTtm, finalModelRow, getColumnIndexByName("EPS TTM"));
                                tableModel.setValueAt(roeTtm, finalModelRow, getColumnIndexByName("ROE TTM"));
                                tableModel.setValueAt(aScore, finalModelRow, getColumnIndexByName("A-Score"));
                                tableModel.setValueAt(epsCurrentYear, finalModelRow, getColumnIndexByName(Watchlist.this.dynamicColumnNames[0]));
                                tableModel.setValueAt(epsNextYear, finalModelRow, getColumnIndexByName(Watchlist.this.dynamicColumnNames[1]));
                                tableModel.setValueAt(epsYear3, finalModelRow, getColumnIndexByName(Watchlist.this.dynamicColumnNames[2]));
                                tableModel.setValueAt(debtToEquityVal, finalModelRow, getColumnIndexByName("Debt to Equity"));
                                tableModel.setValueAt(epsGrowth1, finalModelRow, getColumnIndexByName("EPS Growth 1"));
                                tableModel.setValueAt(currentRatioVal, finalModelRow, getColumnIndexByName("Current Ratio"));
                                tableModel.setValueAt(quickRatioVal, finalModelRow, getColumnIndexByName("Quick Ratio"));
                                tableModel.setValueAt(epsGrowth2, finalModelRow, getColumnIndexByName("EPS Growth 2"));
                                tableModel.setValueAt(epsGrowth3, finalModelRow, getColumnIndexByName("EPS Growth 3"));
                                tableModel.setValueAt(deAvg, finalModelRow, getColumnIndexByName("DE Avg"));
                                tableModel.setValueAt(industry, finalModelRow, getColumnIndexByName("Industry"));
                                tableModel.setValueAt(roeAvg, finalModelRow, getColumnIndexByName("ROE Avg"));
                                tableModel.setValueAt(priceToFCF_TTM, finalModelRow, getColumnIndexByName("P/FCF"));
                                tableModel.setValueAt(PriceToFCF_Avg, finalModelRow, getColumnIndexByName("PFCF Avg"));
                                tableModel.setValueAt(peForward1, finalModelRow, getColumnIndexByName(Watchlist.this.peForwardColumnNames[0]));
                                tableModel.setValueAt(peForward2, finalModelRow, getColumnIndexByName(Watchlist.this.peForwardColumnNames[1]));
                                tableModel.setValueAt(peForward3, finalModelRow, getColumnIndexByName(Watchlist.this.peForwardColumnNames[2]));
                                tableModel.setValueAt(volatilityScore, finalModelRow, getColumnIndexByName("Volatility"));
                                tableModel.setValueAt(prAvg, finalModelRow, getColumnIndexByName("PR Avg"));
                            });
                             workerInternalLastProcessedIndex = i; 
                             System.out.println("Refreshed stock data: " + ticker + " at view index " + i);
                        } else {
                             System.err.println("Failed to fetch stock data for " + ticker + ". stockData is null.");
                        }
                        Thread.sleep(200); 

                    } catch (InterruptedException ex) {
                        Thread.currentThread().interrupt();
                        Watchlist.this.lastSuccessfullyProcessedIndex = workerInternalLastProcessedIndex;
                        System.out.println("Refresh interrupted. Last successfully processed view index: " + Watchlist.this.lastSuccessfullyProcessedIndex);
                        break; 
                    } catch (Exception e) {
                        e.printStackTrace();
                        System.err.println("Error refreshing " + ticker + ": " + e.getMessage());
                    }
                }
                if (!isCancelled()) {
                    Watchlist.this.lastSuccessfullyProcessedIndex = workerInternalLastProcessedIndex;
                }
                return null;
            }

            @Override
            protected void done() {
                SwingUtilities.invokeLater(() -> {
                    contentPane.remove(progressPanel);
                    contentPane.revalidate();
                    contentPane.repaint();

                    try {
                        get(); 
                        if (isCancelled()) {
                            String lastTicker = "N/A";
                            if (Watchlist.this.lastSuccessfullyProcessedIndex != -1 && Watchlist.this.lastSuccessfullyProcessedIndex < tableModel.getRowCount()) {
                                int modelRowForTicker = watchlistTable.convertRowIndexToModel(Watchlist.this.lastSuccessfullyProcessedIndex);
                                if (modelRowForTicker >= 0 && modelRowForTicker < tableModel.getRowCount()) {
                                    lastTicker = (String) tableModel.getValueAt(modelRowForTicker, getColumnIndexByName("Ticker"));
                                }
                            }
                            JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(),
                                    "Refresh Cancelled. Last successfully processed: " + lastTicker + " (view index " + Watchlist.this.lastSuccessfullyProcessedIndex + ").",
                                    "Refresh Interrupted", JOptionPane.WARNING_MESSAGE);
                        } else {
                            if (Watchlist.this.lastSuccessfullyProcessedIndex == tableModel.getRowCount() - 1) {
                                saveWatchlist(); 
                                JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Watchlist refresh completed!", "Refresh Complete", JOptionPane.INFORMATION_MESSAGE);
                                Watchlist.this.lastSuccessfullyProcessedIndex = -1; 
                            } else if (totalItemsForThisRunGlobal <= 0 && tableModel.getRowCount() > 0){ 
                                JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Watchlist is already up-to-date or nothing to resume.", "No Refresh Needed", JOptionPane.INFORMATION_MESSAGE);
                                if(Watchlist.this.lastSuccessfullyProcessedIndex == tableModel.getRowCount() -1) Watchlist.this.lastSuccessfullyProcessedIndex = -1; 
                            } else if (tableModel.getRowCount() == 0) {
                                // Already handled
                            }
                            else if (Watchlist.this.lastSuccessfullyProcessedIndex >= -1) { 
                                 String lastTickerMsg = "None";
                                 if (Watchlist.this.lastSuccessfullyProcessedIndex != -1 && Watchlist.this.lastSuccessfullyProcessedIndex < tableModel.getRowCount()) {
                                     int modelRowForTickerMsg = watchlistTable.convertRowIndexToModel(Watchlist.this.lastSuccessfullyProcessedIndex);
                                      if (modelRowForTickerMsg >= 0 && modelRowForTickerMsg < tableModel.getRowCount()) {
                                        lastTickerMsg = (String) tableModel.getValueAt(modelRowForTickerMsg, getColumnIndexByName("Ticker"));
                                      }
                                 }
                                JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(),
                                    "Refresh finished. Last processed: " + lastTickerMsg + " (view index " + Watchlist.this.lastSuccessfullyProcessedIndex +")." +
                                    (Watchlist.this.lastSuccessfullyProcessedIndex < tableModel.getRowCount() -1 ? " You can resume." : ""),
                                    "Refresh Incomplete/Finished", JOptionPane.INFORMATION_MESSAGE);
                            }
                        }
                    } catch (InterruptedException e) {
                        Thread.currentThread().interrupt();
                        JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Refresh was interrupted.", "Interrupted", JOptionPane.WARNING_MESSAGE);
                    } catch (java.util.concurrent.ExecutionException e) {
                        e.getCause().printStackTrace();
                        JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Error during refresh: " + e.getCause().getMessage(), "Refresh Error", JOptionPane.ERROR_MESSAGE);
                    } finally {
                        refreshDataWorker = null; 
                        saveRefreshState(); // Save the final state
                        updateRefreshButtonsState(); 
                    }
                });
            }
        };
        this.refreshDataWorker.execute();
    }

    private void toggleColumnVisibility(String columnName, boolean visible) {
        TableColumnModel columnModel = watchlistTable.getColumnModel();
        for (int i = 0; i < columnModel.getColumnCount(); i++) {
            TableColumn column = columnModel.getColumn(i);
            if (column.getHeaderValue().equals(columnName)) {
                if (visible) {
                    column.setMinWidth(15);
                    column.setMaxWidth(Integer.MAX_VALUE);
                    column.setPreferredWidth(75); 
                } else {
                    column.setMinWidth(0);
                    column.setMaxWidth(0);
                    column.setPreferredWidth(0);
                }
                break;
            }
        }
    }

    private void loadWatchlist() {
        File file = new File("watchlist.json");
        System.out.println("Attempting to load watchlist from: " + file.getAbsolutePath());
        if (file.exists()) {
            try (FileReader reader = new FileReader(file); Scanner scanner = new Scanner(reader)) {
                String jsonContent = scanner.useDelimiter("\\Z").next();
                JSONArray jsonArray = new JSONArray(jsonContent);
                System.out.println("Found " + jsonArray.length() + " stocks");
                tableModel.setRowCount(0); 
                for (int i = 0; i < jsonArray.length(); i++) {
                    JSONObject jsonObject = jsonArray.getJSONObject(i);
                    Object[] rowData = new Object[tableModel.getColumnCount()]; 
                    
                    rowData[getColumnIndexByName("Name")] = jsonObject.optString("name", "");
                    rowData[getColumnIndexByName("Ticker")] = jsonObject.optString("ticker", "").toUpperCase();
                    rowData[getColumnIndexByName("Refresh Date")] = jsonObject.optString("refreshDate", "N/A");
                    rowData[getColumnIndexByName("Price")] = jsonObject.optDouble("price", 0.0);
                    rowData[getColumnIndexByName("PE TTM")] = jsonObject.optDouble("peTtm", 0.0);
                    rowData[getColumnIndexByName("PB TTM")] = jsonObject.optDouble("pbTtm", 0.0);
                    rowData[getColumnIndexByName("Div. yield %")] = jsonObject.optDouble("dividendYield", 0.0);
                    rowData[getColumnIndexByName("Payout Ratio")] = jsonObject.optDouble("payoutRatio", 0.0);
                    rowData[getColumnIndexByName("Graham Number")] = jsonObject.optDouble("grahamNumber", 0.0);
                    rowData[getColumnIndexByName("PB Avg")] = jsonObject.optDouble("pbAvg", 0.0);
                    rowData[getColumnIndexByName("PE Avg")] = jsonObject.optDouble("peAvg", 0.0);
                    rowData[getColumnIndexByName("EPS TTM")] = jsonObject.optDouble("epsTtm", 0.0);
                    rowData[getColumnIndexByName("ROE TTM")] = jsonObject.optDouble("roeTtm", 0.0);
                    rowData[getColumnIndexByName("A-Score")] = jsonObject.optDouble("aScore", 0.0);
                    rowData[getColumnIndexByName(this.dynamicColumnNames[0])] = jsonObject.optDouble("epsCurrentYear", 0.0);
                    rowData[getColumnIndexByName(this.dynamicColumnNames[1])] = jsonObject.optDouble("epsNextYear", 0.0);
                    rowData[getColumnIndexByName(this.dynamicColumnNames[2])] = jsonObject.optDouble("epsYear3", 0.0);
                    rowData[getColumnIndexByName("Debt to Equity")] = jsonObject.optDouble("debtToEquity", 0.0);
                    rowData[getColumnIndexByName("EPS Growth 1")] = jsonObject.optDouble("epsGrowth1", 0.0);
                    rowData[getColumnIndexByName("Current Ratio")] = jsonObject.optDouble("currentRatio", 0.0);
                    rowData[getColumnIndexByName("Quick Ratio")] = jsonObject.optDouble("quickRatio", 0.0);
                    rowData[getColumnIndexByName("EPS Growth 2")] = jsonObject.optDouble("epsGrowth2", 0.0);
                    rowData[getColumnIndexByName("EPS Growth 3")] = jsonObject.optDouble("epsGrowth3", 0.0);
                    rowData[getColumnIndexByName("DE Avg")] = jsonObject.optDouble("deAvg", 0.0);
                    rowData[getColumnIndexByName("Industry")] = jsonObject.optString("industry", "N/A");
                    rowData[getColumnIndexByName("ROE Avg")] = jsonObject.optDouble("roeAvg", 0.0);
                    rowData[getColumnIndexByName("P/FCF")] = jsonObject.optDouble("priceToFCF_TTM", 0.0);
                    rowData[getColumnIndexByName("PFCF Avg")] = jsonObject.optDouble("priceToFCF_Avg", 0.0);
                    rowData[getColumnIndexByName(this.peForwardColumnNames[0])] = jsonObject.optDouble("peForward1", 0.0);
                    rowData[getColumnIndexByName(this.peForwardColumnNames[1])] = jsonObject.optDouble("peForward2", 0.0);
                    rowData[getColumnIndexByName(this.peForwardColumnNames[2])] = jsonObject.optDouble("peForward3", 0.0);
                    rowData[getColumnIndexByName("Volatility")] = jsonObject.optDouble("volatility", 0.0);
                    rowData[getColumnIndexByName("PR Avg")] = jsonObject.optDouble("payoutRatioAvg", 0.0);
                    
                    tableModel.addRow(rowData);
                }
                System.out.println("Watchlist loading completed");
            } catch (Exception e) { 
                e.printStackTrace();
                JOptionPane.showMessageDialog(null, "Error loading watchlist: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        } else {
            System.out.println("Watchlist file does not exist.");
        }
    }

    private void saveWatchlist() {
        System.out.println("Saving watchlist...");
        JSONArray jsonArray = new JSONArray();
        for (int i = 0; i < tableModel.getRowCount(); i++) {
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("name", tableModel.getValueAt(i, getColumnIndexByName("Name")));
            jsonObject.put("ticker", tableModel.getValueAt(i, getColumnIndexByName("Ticker")));
            jsonObject.put("refreshDate", tableModel.getValueAt(i, getColumnIndexByName("Refresh Date")));
            jsonObject.put("price", tableModel.getValueAt(i, getColumnIndexByName("Price")));
            jsonObject.put("peTtm", tableModel.getValueAt(i, getColumnIndexByName("PE TTM")));
            jsonObject.put("pbTtm", tableModel.getValueAt(i, getColumnIndexByName("PB TTM")));
            jsonObject.put("dividendYield", tableModel.getValueAt(i, getColumnIndexByName("Div. yield %")));
            jsonObject.put("payoutRatio", tableModel.getValueAt(i, getColumnIndexByName("Payout Ratio")));
            jsonObject.put("grahamNumber", tableModel.getValueAt(i, getColumnIndexByName("Graham Number")));
            jsonObject.put("pbAvg", tableModel.getValueAt(i, getColumnIndexByName("PB Avg")));
            jsonObject.put("peAvg", tableModel.getValueAt(i, getColumnIndexByName("PE Avg")));
            jsonObject.put("epsTtm", tableModel.getValueAt(i, getColumnIndexByName("EPS TTM")));
            jsonObject.put("roeTtm", tableModel.getValueAt(i, getColumnIndexByName("ROE TTM")));
            jsonObject.put("aScore", tableModel.getValueAt(i, getColumnIndexByName("A-Score")));
            jsonObject.put("epsCurrentYear", tableModel.getValueAt(i, getColumnIndexByName(this.dynamicColumnNames[0])));
            jsonObject.put("epsNextYear", tableModel.getValueAt(i, getColumnIndexByName(this.dynamicColumnNames[1])));
            jsonObject.put("epsYear3", tableModel.getValueAt(i, getColumnIndexByName(this.dynamicColumnNames[2])));
            jsonObject.put("debtToEquity", tableModel.getValueAt(i, getColumnIndexByName("Debt to Equity")));
            jsonObject.put("epsGrowth1", tableModel.getValueAt(i, getColumnIndexByName("EPS Growth 1")));
            jsonObject.put("currentRatio", tableModel.getValueAt(i, getColumnIndexByName("Current Ratio")));
            jsonObject.put("quickRatio", tableModel.getValueAt(i, getColumnIndexByName("Quick Ratio")));
            jsonObject.put("epsGrowth2", tableModel.getValueAt(i, getColumnIndexByName("EPS Growth 2")));
            jsonObject.put("epsGrowth3", tableModel.getValueAt(i, getColumnIndexByName("EPS Growth 3")));
            jsonObject.put("deAvg", tableModel.getValueAt(i, getColumnIndexByName("DE Avg")));
            jsonObject.put("industry", tableModel.getValueAt(i, getColumnIndexByName("Industry")));
            jsonObject.put("roeAvg", tableModel.getValueAt(i, getColumnIndexByName("ROE Avg")));
            jsonObject.put("priceToFCF_TTM", tableModel.getValueAt(i, getColumnIndexByName("P/FCF")));
            jsonObject.put("priceToFCF_Avg", tableModel.getValueAt(i, getColumnIndexByName("PFCF Avg")));
            jsonObject.put("peForward1", tableModel.getValueAt(i, getColumnIndexByName(this.peForwardColumnNames[0])));
            jsonObject.put("peForward2", tableModel.getValueAt(i, getColumnIndexByName(this.peForwardColumnNames[1])));
            jsonObject.put("peForward3", tableModel.getValueAt(i, getColumnIndexByName(this.peForwardColumnNames[2])));
            jsonObject.put("volatility", tableModel.getValueAt(i, getColumnIndexByName("Volatility")));
            jsonObject.put("payoutRatioAvg", tableModel.getValueAt(i, getColumnIndexByName("PR Avg")));
            jsonArray.put(jsonObject);
        }
        try (FileWriter fileWriter = new FileWriter("watchlist.json")) {
            fileWriter.write(jsonArray.toString(2)); 
            System.out.println("Watchlist saved successfully");
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Error saving watchlist: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

   private void saveColumnSettings() {
        Properties properties = new Properties();
        TableColumnModel columnModel = watchlistTable.getColumnModel();
        Map<Integer, Integer> modelToViewIndexMap = new HashMap<>();
        for (int viewIdx = 0; viewIdx < columnModel.getColumnCount(); viewIdx++) {
            modelToViewIndexMap.put(columnModel.getColumn(viewIdx).getModelIndex(), viewIdx);
        }

        for (int modelIdx = 0; modelIdx < tableModel.getColumnCount(); modelIdx++) {
            TableColumn column = watchlistTable.getColumn(tableModel.getColumnName(modelIdx)); 
            String columnName = column.getHeaderValue().toString();
            
            properties.setProperty("column" + modelIdx + ".name", columnName);
            properties.setProperty("column" + modelIdx + ".modelIndex", String.valueOf(column.getModelIndex()));
            properties.setProperty("column" + modelIdx + ".viewIndex", String.valueOf(modelToViewIndexMap.getOrDefault(column.getModelIndex(), modelIdx)));
            properties.setProperty("column" + modelIdx + ".visible", String.valueOf(column.getMaxWidth() != 0));
            properties.setProperty("column" + modelIdx + ".width", String.valueOf(column.getPreferredWidth()));
        }
        try (FileOutputStream out = new FileOutputStream(COLUMN_SETTINGS_FILE)) {
            properties.store(out, "Column Order, Visibility, and Width Settings");
            System.out.println("Column settings saved successfully");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void loadColumnSettings() {
        Properties properties = new Properties();
        File settingsFile = new File(COLUMN_SETTINGS_FILE);

        if (!settingsFile.exists()) {
            System.out.println("Column settings file not found. Using default column settings.");
            TableColumnModel currentColumnModel = watchlistTable.getColumnModel();
            for (int i = 0; i < currentColumnModel.getColumnCount(); i++) {
                TableColumn col = currentColumnModel.getColumn(i);
                String colName = col.getHeaderValue().toString();
                // Set default visibility for checkboxes
                for (Component comp : columnControlPanel.getComponents()) {
                    if (comp instanceof JCheckBox && ((JCheckBox) comp).getText().equals(colName)) {
                        ((JCheckBox) comp).setSelected(true); // Default to visible
                        break;
                    }
                }
                // Set default column properties (already done by JTable, but good for consistency)
                col.setMinWidth(15);
                col.setMaxWidth(Integer.MAX_VALUE);
                col.setPreferredWidth(75);
            }
            watchlistTable.revalidate();
            watchlistTable.repaint();
            return;
        }

        try (FileInputStream in = new FileInputStream(settingsFile)) {
            properties.load(in);
            TableColumnModel currentColumnModel = watchlistTable.getColumnModel();

            Map<String, Integer> columnToSavedViewIndex = new HashMap<>();
            Map<String, Integer> columnToSavedWidth = new HashMap<>();
            Map<String, Boolean> columnToSavedVisibility = new HashMap<>();
            
            for (int i = 0; ; i++) { 
                String nameKey = "column" + i + ".name"; 
                String propColumnName = properties.getProperty(nameKey);
                if (propColumnName == null) break; 

                int savedViewIndex = Integer.parseInt(properties.getProperty("column" + i + ".viewIndex", String.valueOf(i)));
                boolean isVisible = Boolean.parseBoolean(properties.getProperty("column" + i + ".visible", "true"));
                int savedWidth = Integer.parseInt(properties.getProperty("column" + i + ".width", "75"));

                columnToSavedViewIndex.put(propColumnName, savedViewIndex);
                columnToSavedWidth.put(propColumnName, savedWidth);
                columnToSavedVisibility.put(propColumnName, isVisible);
            }

            List<TableColumn> columnsInViewOrder = new ArrayList<>();
            for (int modelIdx = 0; modelIdx < tableModel.getColumnCount(); modelIdx++) {
                String columnNameFromModel = tableModel.getColumnName(modelIdx);
                TableColumn tableColumn = null;
                try {
                    tableColumn = watchlistTable.getColumn(columnNameFromModel);
                } catch (IllegalArgumentException e) {
                    System.err.println("Could not find TableColumn for: " + columnNameFromModel + " during settings load.");
                    continue; 
                }

                boolean isVisible = columnToSavedVisibility.getOrDefault(columnNameFromModel, true);
                int width = columnToSavedWidth.getOrDefault(columnNameFromModel, 75);

                if (!isVisible) {
                    tableColumn.setMinWidth(0);
                    tableColumn.setMaxWidth(0);
                    tableColumn.setPreferredWidth(0);
                } else {
                    tableColumn.setMinWidth(15);
                    tableColumn.setMaxWidth(Integer.MAX_VALUE);
                    tableColumn.setPreferredWidth(width);
                }
                columnsInViewOrder.add(tableColumn); 

                for (Component comp : columnControlPanel.getComponents()) {
                    if (comp instanceof JCheckBox && ((JCheckBox) comp).getText().equals(columnNameFromModel)) {
                        ((JCheckBox) comp).setSelected(isVisible);
                        break;
                    }
                }
            }
            
            columnsInViewOrder.sort(Comparator.comparingInt(col -> 
                columnToSavedViewIndex.getOrDefault(col.getHeaderValue().toString(), Integer.MAX_VALUE)
            ));

            while (currentColumnModel.getColumnCount() > 0) {
                currentColumnModel.removeColumn(currentColumnModel.getColumn(0));
            }
            for (TableColumn tc : columnsInViewOrder) {
                currentColumnModel.addColumn(tc);
            }
            
            System.out.println("Column settings loaded and applied.");

        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Error loading column settings: " + e.getMessage() + ". Using defaults.", "Settings Error", JOptionPane.ERROR_MESSAGE);
            TableColumnModel currentColumnModel = watchlistTable.getColumnModel();
             for (int i = 0; i < currentColumnModel.getColumnCount(); i++) {
                TableColumn col = currentColumnModel.getColumn(i);
                col.setMinWidth(15);
                col.setMaxWidth(Integer.MAX_VALUE);
                col.setPreferredWidth(75);
                String colName = col.getHeaderValue().toString();
                for (Component comp : columnControlPanel.getComponents()) {
                    if (comp instanceof JCheckBox && ((JCheckBox) comp).getText().equals(colName)) {
                        ((JCheckBox) comp).setSelected(true);
                        break;
                    }
                }
            }
        }
        watchlistTable.revalidate();
        watchlistTable.repaint();
    }


    private void saveSortOrder(java.util.List<? extends RowSorter.SortKey> sortKeys) {
        Properties properties = new Properties();
        for (int i = 0; i < sortKeys.size(); i++) {
            RowSorter.SortKey sortKey = sortKeys.get(i);
            properties.setProperty("sortKey" + i + ".column", String.valueOf(sortKey.getColumn()));
            properties.setProperty("sortKey" + i + ".order", sortKey.getSortOrder().toString());
        }
        try (FileOutputStream out = new FileOutputStream("sort_order.properties")) {
            properties.store(out, "Table Sort Order");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void loadSortOrder(TableRowSorter<DefaultTableModel> sorter) {
        Properties properties = new Properties();
        File file = new File("sort_order.properties");
        if (file.exists()) {
            try (FileInputStream in = new FileInputStream(file)) {
                properties.load(in);
                List<RowSorter.SortKey> sortKeys = new ArrayList<>();
                for (int i = 0; ; i++) {
                    String columnStr = properties.getProperty("sortKey" + i + ".column");
                    String orderStr = properties.getProperty("sortKey" + i + ".order");
                    if (columnStr == null || orderStr == null) break;
                    int column = Integer.parseInt(columnStr);
                    SortOrder order = SortOrder.valueOf(orderStr);
                    if (column < tableModel.getColumnCount()) { 
                       sortKeys.add(new RowSorter.SortKey(column, order));
                    }
                }
                if (!sortKeys.isEmpty()) sorter.setSortKeys(sortKeys);
            } catch (Exception e) { 
                e.printStackTrace();
            }
        }
    }

    private void addStock() {
        String ticker = JOptionPane.showInputDialog(watchlistTable.getTopLevelAncestor(), "Enter Stock Ticker:");
        if (ticker != null && !ticker.trim().isEmpty()) {
            ticker = ticker.toUpperCase();
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                if (ticker.equals(tableModel.getValueAt(i, getColumnIndexByName("Ticker")))) {
                    JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Stock '" + ticker + "' is already in the watchlist.", "Duplicate Stock", JOptionPane.WARNING_MESSAGE);
                    return;
                }
            }
            try {
                JSONObject stockData = fetchStockData(ticker);
                if (stockData == null || !stockData.has("name")) { 
                    JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Symbol '" + ticker + "' not found or no data available.", "Symbol Not Found", JOptionPane.WARNING_MESSAGE);
                    return;
                }
                String name = stockData.getString("name");
                Object[] rowData = new Object[tableModel.getColumnCount()];
                rowData[getColumnIndexByName("Name")] = name;
                rowData[getColumnIndexByName("Ticker")] = ticker;
                rowData[getColumnIndexByName("Refresh Date")] = "N/A"; 

                for(int j=0; j<rowData.length; j++) { 
                    if (rowData[j] == null) { 
                        if(tableModel.getColumnClass(j) == Double.class) rowData[j] = 0.0;
                        else if (tableModel.getColumnName(j).equals("Industry")) rowData[j] = "N/A"; 
                        else rowData[j] = ""; 
                    }
                }
                tableModel.addRow(rowData);
                System.out.println("Added stock: " + ticker);
                saveWatchlist(); 
            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Error processing '" + ticker + "': " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }
    
    private void refreshSingleStock(String ticker, int modelRow) {
        System.out.println("Starting single stock refresh for: " + ticker);
        JDialog progressDialog = new JDialog((Frame) watchlistTable.getTopLevelAncestor(), "Refreshing " + ticker, false); 
        progressDialog.setSize(350, 120);
        progressDialog.setLayout(new BorderLayout());
        progressDialog.setLocationRelativeTo(watchlistTable.getTopLevelAncestor());
        
        JProgressBar progressBarSingle = new JProgressBar();
        progressBarSingle.setIndeterminate(true);
        JLabel statusLabelSingle = new JLabel(" Fetching data for " + ticker + "...", SwingConstants.CENTER);
        statusLabelSingle.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        
        JPanel singleProgressPanel = new JPanel(new BorderLayout());
        singleProgressPanel.add(statusLabelSingle, BorderLayout.NORTH);
        singleProgressPanel.add(progressBarSingle, BorderLayout.CENTER);
        progressDialog.add(singleProgressPanel);

        SwingWorker<Boolean, Void> singleStockWorker = new SwingWorker<Boolean, Void>() {
            @Override
            protected Boolean doInBackground() {
                try {
                    JSONObject stockData = fetchStockData(ticker);
                    JSONObject ratios = fetchStockRatios(ticker);
                    JSONObject epsEstimates = Estimates.fetchEpsEstimates(ticker);

                    if (stockData != null) {
                        final double price = convertGbxToGbp(round(stockData.optDouble("price",0.0), 2), ticker);
                        double peTtm = round(ratios.optDouble("peRatioTTM", 0.0), 2);
                        double pbTtm = round(ratios.optDouble("pbRatioTTM", 0.0), 2);
                        double epsTtm = peTtm != 0 ? round((1 / peTtm) * price, 2) : 0.0;
                        double roeTtm = round(ratios.optDouble("roeTTM", 0.0), 2);
                        double dividendYieldTTM = KeyMetricsTTM.fetchDividendYield(ticker);
                        double payoutRatioVal = KeyMetricsTTM.fetchPayoutRatio(ticker);
                        double debtToEquityVal = KeyMetricsTTM.fetchDebtToEquity(ticker);
                        double epsCurrentYear = epsEstimates != null ? round(epsEstimates.optDouble("eps0", 0.0), 2) : 0.0;
                        double epsNextYear = epsEstimates != null ? round(epsEstimates.optDouble("eps1", 0.0), 2) : 0.0;
                        double epsYear3 = epsEstimates != null ? round(epsEstimates.optDouble("eps2", 0.0), 2) : 0.0;
                        double currentRatioVal = KeyMetricsTTM.fetchCurrentRatio(ticker);
                        double quickRatioVal = 0.0; 
                        double epsGrowth1 = calculateEpsGrowth1(epsCurrentYear, epsTtm);
                        double epsGrowth2 = calculateEpsGrowth2(epsCurrentYear, epsNextYear);
                        double epsGrowth3 = calculateEpsGrowth3(epsYear3, epsNextYear);
                        double pbAvg = fetchAveragePB(ticker);
                        double peForward1 = calculatePEForward(price, epsCurrentYear);
                        double peForward2 = calculatePEForward(price, epsNextYear);
                        double peForward3 = calculatePEForward(price, epsYear3);
                        double peAvg = fetchAveragePE(ticker);
                        double priceToFCF_TTM = TTM_Ratios.getPriceToFreeCashFlowRatioTTM(ticker);
                        double PriceToFCF_Avg = Ratios.fetchPriceToFreeCashFlowAverage(ticker);
                        double roeAvg = fetchAverageROE(ticker);
                        double grahamNumber = calculateGrahamNumber(price, peAvg, pbAvg, epsTtm, pbTtm);
                        double deAvg = Ratios.fetchDebtToEquityAverage(ticker);
                        String industry = CompanyOverview.fetchIndustry(ticker);
                        double prAvg = Ratios.fetchPayoutRatioAverage(ticker);
                        double aScore = calculateAScore(pbAvg, pbTtm, peAvg, peTtm, payoutRatioVal, debtToEquityVal, roeTtm, roeAvg, dividendYieldTTM, deAvg, epsGrowth1, epsGrowth2, epsGrowth3, currentRatioVal, quickRatioVal, grahamNumber, price, priceToFCF_TTM, PriceToFCF_Avg, prAvg);
                        double volatilityScore = Analytics.calculateVolatilityScore(ticker);
                        String refreshDateStr = LocalDate.now().format(DATE_FORMATTER);
                        
                        SwingUtilities.invokeLater(() -> {
                            tableModel.setValueAt(refreshDateStr, modelRow, getColumnIndexByName("Refresh Date"));
                            tableModel.setValueAt(price, modelRow, getColumnIndexByName("Price"));
                            tableModel.setValueAt(peTtm, modelRow, getColumnIndexByName("PE TTM"));
                            tableModel.setValueAt(pbTtm, modelRow, getColumnIndexByName("PB TTM"));
                            tableModel.setValueAt(dividendYieldTTM, modelRow, getColumnIndexByName("Div. yield %"));
                            tableModel.setValueAt(payoutRatioVal, modelRow, getColumnIndexByName("Payout Ratio"));
                            tableModel.setValueAt(grahamNumber, modelRow, getColumnIndexByName("Graham Number"));
                            tableModel.setValueAt(pbAvg, modelRow, getColumnIndexByName("PB Avg"));
                            tableModel.setValueAt(peAvg, modelRow, getColumnIndexByName("PE Avg"));
                            tableModel.setValueAt(epsTtm, modelRow, getColumnIndexByName("EPS TTM"));
                            tableModel.setValueAt(roeTtm, modelRow, getColumnIndexByName("ROE TTM"));
                            tableModel.setValueAt(aScore, modelRow, getColumnIndexByName("A-Score"));
                            tableModel.setValueAt(epsCurrentYear, modelRow, getColumnIndexByName(Watchlist.this.dynamicColumnNames[0]));
                            tableModel.setValueAt(epsNextYear, modelRow, getColumnIndexByName(Watchlist.this.dynamicColumnNames[1]));
                            tableModel.setValueAt(epsYear3, modelRow, getColumnIndexByName(Watchlist.this.dynamicColumnNames[2]));
                            tableModel.setValueAt(debtToEquityVal, modelRow, getColumnIndexByName("Debt to Equity"));
                            tableModel.setValueAt(epsGrowth1, modelRow, getColumnIndexByName("EPS Growth 1"));
                            tableModel.setValueAt(currentRatioVal, modelRow, getColumnIndexByName("Current Ratio"));
                            tableModel.setValueAt(quickRatioVal, modelRow, getColumnIndexByName("Quick Ratio"));
                            tableModel.setValueAt(epsGrowth2, modelRow, getColumnIndexByName("EPS Growth 2"));
                            tableModel.setValueAt(epsGrowth3, modelRow, getColumnIndexByName("EPS Growth 3"));
                            tableModel.setValueAt(deAvg, modelRow, getColumnIndexByName("DE Avg"));
                            tableModel.setValueAt(industry, modelRow, getColumnIndexByName("Industry"));
                            tableModel.setValueAt(roeAvg, modelRow, getColumnIndexByName("ROE Avg"));
                            tableModel.setValueAt(priceToFCF_TTM, modelRow, getColumnIndexByName("P/FCF"));
                            tableModel.setValueAt(PriceToFCF_Avg, modelRow, getColumnIndexByName("PFCF Avg"));
                            tableModel.setValueAt(peForward1, modelRow, getColumnIndexByName(Watchlist.this.peForwardColumnNames[0]));
                            tableModel.setValueAt(peForward2, modelRow, getColumnIndexByName(Watchlist.this.peForwardColumnNames[1]));
                            tableModel.setValueAt(peForward3, modelRow, getColumnIndexByName(Watchlist.this.peForwardColumnNames[2]));
                            tableModel.setValueAt(volatilityScore, modelRow, getColumnIndexByName("Volatility"));
                            tableModel.setValueAt(prAvg, modelRow, getColumnIndexByName("PR Avg"));
                            saveWatchlist(); 
                        });
                        return true; 
                    }
                    return false; 
                } catch (Exception e) {
                    e.printStackTrace();
                    SwingUtilities.invokeLater(() -> JOptionPane.showMessageDialog(progressDialog.getParent(), "Error refreshing " + ticker + ": " + e.getMessage(), "Refresh Error", JOptionPane.ERROR_MESSAGE));
                    return false; 
                }
            }

            @Override
            protected void done() {
                SwingUtilities.invokeLater(() -> {
                    progressDialog.dispose();
                    try {
                        boolean success = get(); 
                        String message = success ? ticker + " data refreshed successfully!" : "Failed to refresh " + ticker + ".";
                        JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), message, "Stock Refresh", 
                                                     success ? JOptionPane.INFORMATION_MESSAGE : JOptionPane.ERROR_MESSAGE);
                    } catch (InterruptedException e) {
                        Thread.currentThread().interrupt();
                         JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), ticker + " refresh was interrupted.", "Interrupted", JOptionPane.WARNING_MESSAGE);
                    } catch (java.util.concurrent.ExecutionException e) {
                         JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Error during " + ticker + " refresh: " + e.getCause().getMessage(), "Refresh Error", JOptionPane.ERROR_MESSAGE);
                    }
                });
            }
        };
        SwingUtilities.invokeLater(() -> progressDialog.setVisible(true)); 
        singleStockWorker.execute();
    }

    private void clearStockValues(int modelRow) {
        String ticker = (String) tableModel.getValueAt(modelRow, getColumnIndexByName("Ticker"));
        int confirm = JOptionPane.showConfirmDialog(
                watchlistTable.getTopLevelAncestor(),
                "Are you sure you want to clear all values for " + ticker + "?",
                "Confirm Clear",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.QUESTION_MESSAGE);

        if (confirm == JOptionPane.YES_OPTION) {
            for (int col = 0; col < tableModel.getColumnCount(); col++) { 
                if (col == getColumnIndexByName("Name") || col == getColumnIndexByName("Ticker")) {
                    continue; 
                }
                if (tableModel.getColumnClass(col) == Double.class) {
                    tableModel.setValueAt(0.0, modelRow, col);
                } else if (tableModel.getColumnName(col).equals("Industry") || tableModel.getColumnName(col).equals("Refresh Date")) { 
                    tableModel.setValueAt("N/A", modelRow, col);
                }
            }
            saveWatchlist();
            JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Values for " + ticker + " have been cleared.", "Clear Complete", JOptionPane.INFORMATION_MESSAGE);
        }
    }
    
    private void deleteStock() {
        int selectedRow = watchlistTable.getSelectedRow();
        if (selectedRow != -1) {
            int modelRow = watchlistTable.convertRowIndexToModel(selectedRow);
            String ticker = (String) tableModel.getValueAt(modelRow, getColumnIndexByName("Ticker"));
            int confirm = JOptionPane.showConfirmDialog(
                    watchlistTable.getTopLevelAncestor(),
                    "Are you sure you want to delete " + ticker + " from the watchlist?",
                    "Confirm Delete",
                    JOptionPane.YES_NO_OPTION,
                    JOptionPane.WARNING_MESSAGE);
            if (confirm == JOptionPane.YES_OPTION) {
                tableModel.removeRow(modelRow);
                saveWatchlist();
                System.out.println("Stock deleted successfully: " + ticker);
                if (lastSuccessfullyProcessedIndex >= modelRow && lastSuccessfullyProcessedIndex > -1) { 
                    lastSuccessfullyProcessedIndex--; 
                } else if (modelRow < lastSuccessfullyProcessedIndex) {
                    // This condition logic might need review based on whether lastSuccessfullyProcessedIndex refers to view or model index
                    // Currently, it's intended to be a view index before conversion inside the worker.
                }
                updateRefreshButtonsState();
            }
        } else {
            JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Please select a stock to delete.", "No Selection", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private double round(double value, int places) {
        if (places < 0) throw new IllegalArgumentException();
        BigDecimal bd = BigDecimal.valueOf(value);
        bd = bd.setScale(places, RoundingMode.HALF_UP);
        return bd.doubleValue();
    }

    private JSONObject fetchStockData(String ticker) {
        String urlString = String.format("https://financialmodelingprep.com/api/v3/quote/%s?apikey=%s", ticker, API_KEY);
        HttpURLConnection connection = null;
        try {
            URL url = new URL(urlString);
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(5000);
            int responseCode = connection.getResponseCode();
            if (responseCode == 200) {
                try (Scanner scanner = new Scanner(connection.getInputStream())) {
                    String response = scanner.useDelimiter("\\A").next();
                    JSONArray data = new JSONArray(response);
                    return data.length() > 0 ? data.getJSONObject(0) : null;
                }
            } else {
                 System.err.println("fetchStockData HTTP Error: " + responseCode + " for " + ticker);
            }
        } catch (Exception e) {
            System.err.println("Exception in fetchStockData for " + ticker + ": " + e.getMessage());
            e.printStackTrace(); 
        } finally {
            if (connection != null) connection.disconnect();
        }
        return null;
    }

    private JSONObject fetchStockRatios(String ticker) {
        JSONObject combinedRatios = new JSONObject();
        String urlMetrics = String.format("https://financialmodelingprep.com/api/v3/key-metrics-ttm/%s?apikey=%s", ticker, API_KEY);
        String urlRatios = String.format("https://financialmodelingprep.com/api/v3/ratios-ttm/%s?apikey=%s", ticker, API_KEY);

        HttpURLConnection conn1 = null;
        try {
            URL url1 = new URL(urlMetrics);
            conn1 = (HttpURLConnection) url1.openConnection();
            conn1.setRequestMethod("GET");
            conn1.setConnectTimeout(5000); 
            conn1.setReadTimeout(5000);
            if (conn1.getResponseCode() == 200) {
                try (Scanner scanner1 = new Scanner(conn1.getInputStream())) {
                    String response1 = scanner1.useDelimiter("\\A").next();
                    JSONArray data1 = new JSONArray(response1);
                    if (data1.length() > 0) {
                        JSONObject metricsData = data1.getJSONObject(0);
                        for (String key : metricsData.keySet()) combinedRatios.put(key, metricsData.get(key));
                    }
                }
            } else {
                 System.err.println("fetchStockRatios (key-metrics-ttm) HTTP Error: " + conn1.getResponseCode() + " for " + ticker);
            }
        } catch (Exception e) {
             System.err.println("Exception in fetchStockRatios (key-metrics-ttm) for " + ticker + ": " + e.getMessage());
             e.printStackTrace();
        } finally {
            if (conn1 != null) conn1.disconnect();
        }

        HttpURLConnection conn2 = null;
        try {
            URL url2 = new URL(urlRatios);
            conn2 = (HttpURLConnection) url2.openConnection();
            conn2.setRequestMethod("GET");
            conn2.setConnectTimeout(5000);
            conn2.setReadTimeout(5000);
            if (conn2.getResponseCode() == 200) {
                 try (Scanner scanner2 = new Scanner(conn2.getInputStream())) {
                    String response2 = scanner2.useDelimiter("\\A").next();
                    JSONArray data2 = new JSONArray(response2);
                    if (data2.length() > 0) {
                        JSONObject ratiosData = data2.getJSONObject(0);
                        for (String key : ratiosData.keySet()) {
                            if (!combinedRatios.has(key) || (combinedRatios.optDouble(key, Double.NaN) == 0.0 && ratiosData.optDouble(key, Double.NaN) != 0.0) ) { 
                                combinedRatios.put(key, ratiosData.get(key));
                            }
                        }
                    }
                }
            } else {
                 System.err.println("fetchStockRatios (ratios-ttm) HTTP Error: " + conn2.getResponseCode() + " for " + ticker);
            }
        } catch (Exception e) {
            System.err.println("Exception in fetchStockRatios (ratios-ttm) for " + ticker + ": " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (conn2 != null) conn2.disconnect();
        }
        return combinedRatios;
    }
    
    private double fetchAveragePB(String ticker) {
        String urlString = String.format("https://financialmodelingprep.com/api/v3/key-metrics/%s?period=annual&limit=20&apikey=%s", ticker, API_KEY);
        HttpURLConnection connection = null;
        double sum = 0; int count = 0;
        try {
            URL url = new URL(urlString);
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(5000); connection.setReadTimeout(5000);
            if (connection.getResponseCode() == 200) {
                try(Scanner scanner = new Scanner(connection.getInputStream())) {
                    String response = scanner.useDelimiter("\\A").next();
                    if (response.trim().isEmpty() || response.trim().equals("[]")) return 0.0; 
                    JSONArray data = new JSONArray(response);
                    for (int i = 0; i < data.length(); i++) {
                        JSONObject metrics = data.getJSONObject(i);
                        double pbRatio = metrics.optDouble("pbRatio", Double.NaN); 
                        if (!Double.isNaN(pbRatio) && pbRatio != 0.0) { 
                            pbRatio = Math.max(-10.0, Math.min(pbRatio, 10.0)); 
                            sum += pbRatio; count++;
                        }
                    }
                }
            }
        } catch (Exception e) { e.printStackTrace(); } 
        finally { if (connection != null) connection.disconnect(); }
        return count > 0 ? round(sum / count, 2) : 0.0;
    }

    private double fetchAveragePE(String ticker) {
        String urlString = String.format("https://financialmodelingprep.com/api/v3/key-metrics/%s?period=annual&limit=20&apikey=%s", ticker, API_KEY);
        HttpURLConnection connection = null;
        double sum = 0; int count = 0;
        try {
            URL url = new URL(urlString);
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(5000); connection.setReadTimeout(5000);
            if (connection.getResponseCode() == 200) {
                try(Scanner scanner = new Scanner(connection.getInputStream())) {
                    String response = scanner.useDelimiter("\\A").next();
                     if (response.trim().isEmpty() || response.trim().equals("[]")) return 0.0;
                    JSONArray data = new JSONArray(response);
                    for (int i = 0; i < data.length(); i++) {
                        JSONObject metrics = data.getJSONObject(i);
                        double peRatio = metrics.optDouble("peRatio", Double.NaN);
                        if (!Double.isNaN(peRatio) && peRatio > 0.0) { 
                            peRatio = Math.min(peRatio, 40.0); 
                            sum += peRatio; count++;
                        }
                    }
                }
            }
        } catch (Exception e) { e.printStackTrace(); }
        finally { if (connection != null) connection.disconnect(); }
        return count > 0 ? round(sum / count, 2) : 0.0;
    }
    
    private double fetchAverageROE(String ticker) {
        String urlString = String.format("https://financialmodelingprep.com/api/v3/ratios/%s?period=annual&limit=20&apikey=%s", ticker, API_KEY);
        HttpURLConnection connection = null;
        double sum = 0; int count = 0;
        try {
            URL url = new URL(urlString);
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(5000); connection.setReadTimeout(5000);
            if (connection.getResponseCode() == 200) {
                try(Scanner scanner = new Scanner(connection.getInputStream())) {
                    String response = scanner.useDelimiter("\\A").next();
                    if (response.trim().isEmpty() || response.trim().equals("[]")) return 0.0;
                    JSONArray data = new JSONArray(response);
                    for (int i = 0; i < data.length(); i++) {
                        JSONObject metrics = data.getJSONObject(i);
                        double roe = metrics.optDouble("returnOnEquity", metrics.optDouble("roe", Double.NaN)); 
                        if (!Double.isNaN(roe) && roe != 0.0) { 
                             roe = Math.max(-0.50, Math.min(roe, 0.50)); 
                            sum += roe; count++;
                        }
                    }
                }
            }
        } catch (Exception e) { e.printStackTrace(); }
        finally { if (connection != null) connection.disconnect(); }
        return count > 0 ? round(sum / count, 2) : 0.0;
    }

    private double calculateGrahamNumber(double price, double peAvg, double pbAvg, double epsTtm, double pbTtm) {
        if (epsTtm <= 0 || price <= 0 || pbTtm <= 0) { 
            return 0.0;
        }
        double bvps = price / pbTtm; 
        double grahamValue = Math.sqrt(22.5 * epsTtm * bvps);
        return round(grahamValue, 2);
    }
    
    private double convertGbxToGbp(double price, String ticker) {
        if (ticker != null && ticker.endsWith(".L")) {
            return round(price / 100.0, 2);
        }
        return price;
    }

    private double calculateEpsGrowth1(double epsCurrentYear, double epsTtm) {
        if (epsTtm == 0) return 0.0;
        if (epsTtm < 0 && epsCurrentYear > 0) return 100.0; 
        if (epsTtm > 0 && epsCurrentYear < 0) return -100.0; 
        return round(100 * (epsCurrentYear - epsTtm) / Math.abs(epsTtm), 2);
    }

    private double calculateEpsGrowth2(double epsCurrentYear, double epsNextYear) {
        if (epsCurrentYear == 0) return 0.0;
         if (epsCurrentYear < 0 && epsNextYear > 0) return 100.0;
        if (epsCurrentYear > 0 && epsNextYear < 0) return -100.0;
        return round(100 * (epsNextYear - epsCurrentYear) / Math.abs(epsCurrentYear), 2);
    }

    private double calculateEpsGrowth3(double epsYear3, double epsNextYear) {
        if (epsNextYear == 0) return 0.0;
        if (epsNextYear < 0 && epsYear3 > 0) return 100.0;
        if (epsNextYear > 0 && epsYear3 < 0) return -100.0;
        return round(100 * (epsYear3 - epsNextYear) / Math.abs(epsNextYear), 2);
    }
    
    private double calculatePEForward(double price, double eps) {
        if (eps <= 0) return 0.0; 
        return round(price / eps, 2);
    }

    // getPEForwardColumnNames was already defined as instance method.

    private double calculateAScore(double pbAvg, double pbTtm, double peAvg, double peTtm, double payoutRatio, double debtToEquity,
                               double roeTtm, double roeAvg, double dividendYieldTTM, double deAvg, double epsGrowth1, double epsGrowth2, double epsGrowth3,
                               double currentRatioVal, double quickRatio, double grahamNumber, double price, double priceToFCF_TTM, double PriceToFCF_Avg, double prAvg) {
        double score = 0;
        if (peTtm > 0 && peAvg > 0) { if (peTtm < peAvg * 0.75) score += 2; else if (peTtm < peAvg) score += 1; } 
        else if (peTtm > 0 && peTtm < 15) score +=1;

        if (pbTtm > 0 && pbAvg > 0) { if (pbTtm < pbAvg * 0.75) score += 2; else if (pbTtm < pbAvg) score += 1; } 
        else if (pbTtm > 0 && pbTtm < 1.5) score +=1;

        if (payoutRatio > 0.05 && payoutRatio <= 0.60) score += 2; 
        else if (payoutRatio > 0.60 && payoutRatio <= 0.80) score += 1;

        if (dividendYieldTTM >= 0.04) score += 2; 
        else if (dividendYieldTTM >= 0.02) score += 1;

        if (debtToEquity >= 0 && deAvg > 0) { if (debtToEquity < deAvg * 0.75) score += 2; else if (debtToEquity < deAvg) score += 1;} 
        else if (debtToEquity >= 0 && debtToEquity < 0.5) score += 2; 
        else if (debtToEquity >= 0 && debtToEquity < 1.0) score +=1;

        if (roeTtm > 0 && roeAvg != 0) { if (roeTtm > roeAvg * 1.25) score += 2; else if (roeTtm > roeAvg) score += 1;} 
        else if (roeTtm > 0.15) score += 2; 
        else if (roeTtm > 0.10) score += 1;

        if (epsGrowth1 > 10) score += 1; else if (epsGrowth1 < -10) score -=1;
        if (epsGrowth2 > 10) score += 1; else if (epsGrowth2 < -10) score -=1;
        if (epsGrowth3 > 10) score += 1; else if (epsGrowth3 < -10) score -=1;
        
        if (currentRatioVal >= 1.5) score += 1;

        if (price > 0 && grahamNumber > 0) { if (grahamNumber > price * 1.25) score += 2; else if (grahamNumber > price) score += 1; }
        
        if (priceToFCF_TTM > 0 && PriceToFCF_Avg > 0) { if (priceToFCF_TTM < PriceToFCF_Avg * 0.75) score += 2; else if (priceToFCF_TTM < PriceToFCF_Avg) score += 1; } 
        else if (priceToFCF_TTM > 0 && priceToFCF_TTM < 15) score +=1;

        if (prAvg > 0 && payoutRatio > 0) { if (payoutRatio < prAvg * 0.75) score += 2; else if (payoutRatio < prAvg) score += 1; }

        return Math.max(0, score); 
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new Watchlist().createAndShowGUI());
    }

    private double getPriceForRow(int modelRow) {
        int priceColumn = getColumnIndexByName("Price"); 
        if (priceColumn != -1 && modelRow >=0 && modelRow < tableModel.getRowCount()) { 
            Object value = tableModel.getValueAt(modelRow, priceColumn);
            if (value instanceof Double) return (Double) value;
        }
        return 0.0;
    }
     private double getPEAvgForRow(int modelRow) {
        int colIdx = getColumnIndexByName("PE Avg"); 
        if (colIdx != -1 && modelRow >=0 && modelRow < tableModel.getRowCount() && tableModel.getValueAt(modelRow, colIdx) instanceof Double) {
            return (Double) tableModel.getValueAt(modelRow, colIdx);
        }
        return 0.0;
    }

    private double getPBAvgForRow(int modelRow) {
        int colIdx = getColumnIndexByName("PB Avg"); 
        if (colIdx != -1 && modelRow >=0 && modelRow < tableModel.getRowCount() && tableModel.getValueAt(modelRow, colIdx) instanceof Double) {
            return (Double) tableModel.getValueAt(modelRow, colIdx);
        }
        return 0.0;
    }
    
    private double getPriceToFCFAvgForRow(int modelRow) {
        int colIdx = getColumnIndexByName("PFCF Avg"); 
        if (colIdx != -1 && modelRow >=0 && modelRow < tableModel.getRowCount() && tableModel.getValueAt(modelRow, colIdx) instanceof Double) {
            return (Double) tableModel.getValueAt(modelRow, colIdx);
        }
        return 0.0;
    }
    
    private double getROEAvgForRow(int modelRow) {
        int colIdx = getColumnIndexByName("ROE Avg"); 
         if (colIdx != -1 && modelRow >=0 && modelRow < tableModel.getRowCount() && tableModel.getValueAt(modelRow, colIdx) instanceof Double) {
            return (Double) tableModel.getValueAt(modelRow, colIdx);
        }
        return 0.0;
    }

    private double getDEAvgForRow(int modelRow) {
        int colIdx = getColumnIndexByName("DE Avg"); 
        if (colIdx != -1 && modelRow >=0 && modelRow < tableModel.getRowCount() && tableModel.getValueAt(modelRow, colIdx) instanceof Double) {
            return (Double) tableModel.getValueAt(modelRow, colIdx);
        }
        return 0.0;
    }


    private void exportToExcel() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save Excel File");
        fileChooser.setSelectedFile(new File("watchlist_export.xlsx"));

        if (fileChooser.showSaveDialog(watchlistTable.getTopLevelAncestor()) == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();
            if (!file.getName().toLowerCase().endsWith(".xlsx")) {
                file = new File(file.getAbsolutePath() + ".xlsx");
            }

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Watchlist");

                XSSFCellStyle headerStyle = workbook.createCellStyle();
                org.apache.poi.ss.usermodel.Font headerFont = workbook.createFont(); 
                headerFont.setBold(true);
                headerStyle.setFont(headerFont);
                headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                headerStyle.setBorderBottom(BorderStyle.THIN);
                
                XSSFCellStyle lightRedStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 255, (byte) 235, (byte) 235});
                XSSFCellStyle mediumRedStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte)255, (byte)200, (byte)200});
                XSSFCellStyle darkRedStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte)255, (byte)180, (byte)180});
                XSSFCellStyle lightYellowStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 255, (byte) 255, (byte) 220});
                XSSFCellStyle lightGreenStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 220, (byte) 255, (byte) 220});
                XSSFCellStyle mediumGreenStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 198, (byte) 255, (byte) 198});
                XSSFCellStyle darkGreenStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 178, (byte) 255, (byte) 178});
                XSSFCellStyle lightPinkStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 255, (byte) 230, (byte) 230});
                XSSFCellStyle mediumPinkStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 255, (byte) 200, (byte) 200});
                XSSFCellStyle darkPinkStyleExcel = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 255, (byte) 180, (byte) 180}); 
                XSSFCellStyle normalStyle = workbook.createCellStyle(); 


                TableColumnModel columnModel = watchlistTable.getColumnModel();
                Row excelHeaderRow = sheet.createRow(0);
                int excelColIdx = 0;
                for (int viewCol = 0; viewCol < watchlistTable.getColumnCount(); viewCol++) {
                    TableColumn tableColumn = columnModel.getColumn(viewCol);
                    if (tableColumn.getMaxWidth() > 0) { 
                        Cell cell = excelHeaderRow.createCell(excelColIdx++); 
                        cell.setCellValue(tableModel.getColumnName(tableColumn.getModelIndex()));
                        cell.setCellStyle(headerStyle);
                    }
                }
                
                DataFormat dataFormat = workbook.createDataFormat();

                for (int tblViewRow = 0; tblViewRow < watchlistTable.getRowCount(); tblViewRow++) { 
                    int tblModelRow = watchlistTable.convertRowIndexToModel(tblViewRow); 
                    Row excelDataRow = sheet.createRow(tblViewRow + 1);
                    excelColIdx = 0; 
                    
                    for (int viewCol = 0; viewCol < watchlistTable.getColumnCount(); viewCol++) {
                        TableColumn tableColumn = columnModel.getColumn(viewCol);
                        if (tableColumn.getMaxWidth() == 0) continue; 

                        int modelCol = tableColumn.getModelIndex();
                        Object value = tableModel.getValueAt(tblModelRow, modelCol); 
                        Cell cell = excelDataRow.createCell(excelColIdx++);
                        String columnName = tableModel.getColumnName(modelCol);

                        if (value instanceof Double) {
                            double numValue = (Double) value;
                            cell.setCellValue(numValue);
                             XSSFCellStyle cellStyle = workbook.createCellStyle(); 
                             cellStyle.cloneStyleFrom(normalStyle); 
                             cellStyle.setDataFormat(dataFormat.getFormat("0.00"));

                            if (columnName.equals("Graham Number")) {
                                double price = getPriceForRow(tblModelRow); 
                                if (price > 0) {
                                    double percentDiff = (numValue - price) / price * 100;
                                    if (percentDiff > 50) cellStyle.setFillForegroundColor(darkGreenStyle.getFillForegroundColorColor());
                                    else if (percentDiff > 25) cellStyle.setFillForegroundColor(mediumGreenStyle.getFillForegroundColorColor());
                                    else if (percentDiff > 0) cellStyle.setFillForegroundColor(lightGreenStyle.getFillForegroundColorColor());
                                    else if (percentDiff <= 0 && percentDiff > -25) cellStyle.setFillForegroundColor(lightPinkStyle.getFillForegroundColorColor()); 
                                    else if (percentDiff <= -25 && percentDiff > -50) cellStyle.setFillForegroundColor(mediumPinkStyle.getFillForegroundColorColor());
                                    else if (percentDiff <= -50) cellStyle.setFillForegroundColor(darkPinkStyleExcel.getFillForegroundColorColor());
                                }
                            } else if (columnName.equals("PE TTM") || columnName.startsWith("PE FWD")) {
                                double peAvg = getPEAvgForRow(tblModelRow); 
                                if (peAvg > 0 && numValue > 0) {
                                    double ratio = numValue / peAvg;
                                    if (ratio < 0.5) cellStyle.setFillForegroundColor(darkGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 0.75) cellStyle.setFillForegroundColor(mediumGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.0) cellStyle.setFillForegroundColor(lightGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.25) cellStyle.setFillForegroundColor(lightPinkStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.5) cellStyle.setFillForegroundColor(mediumPinkStyle.getFillForegroundColorColor());
                                    else cellStyle.setFillForegroundColor(darkPinkStyleExcel.getFillForegroundColorColor());
                                } else if (numValue < 0) cellStyle.setFillForegroundColor(lightRedStyle.getFillForegroundColorColor());
                                else if (numValue == 0) cellStyle.setFillForegroundColor(lightYellowStyle.getFillForegroundColorColor());
                            } else if (columnName.equals("PB TTM")) {
                                double pbAvg = getPBAvgForRow(tblModelRow);
                                 if (pbAvg > 0 && numValue > 0) {
                                    double ratio = numValue / pbAvg;
                                    if (ratio < 0.5) cellStyle.setFillForegroundColor(darkGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 0.75) cellStyle.setFillForegroundColor(mediumGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.0) cellStyle.setFillForegroundColor(lightGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.25) cellStyle.setFillForegroundColor(lightPinkStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.5) cellStyle.setFillForegroundColor(mediumPinkStyle.getFillForegroundColorColor());
                                    else cellStyle.setFillForegroundColor(darkPinkStyleExcel.getFillForegroundColorColor());
                                 } else if (numValue < 0) cellStyle.setFillForegroundColor(lightRedStyle.getFillForegroundColorColor());
                                 else if (numValue == 0) cellStyle.setFillForegroundColor(lightYellowStyle.getFillForegroundColorColor());
                            } else if (columnName.equals("P/FCF")) {
                                double pfcfAvg = getPriceToFCFAvgForRow(tblModelRow);
                                if (pfcfAvg > 0 && numValue > 0) {
                                    double ratio = numValue / pfcfAvg;
                                    if (ratio < 0.5) cellStyle.setFillForegroundColor(darkGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 0.75) cellStyle.setFillForegroundColor(mediumGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.0) cellStyle.setFillForegroundColor(lightGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.25) cellStyle.setFillForegroundColor(lightPinkStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.5) cellStyle.setFillForegroundColor(mediumPinkStyle.getFillForegroundColorColor());
                                    else cellStyle.setFillForegroundColor(darkPinkStyleExcel.getFillForegroundColorColor());
                                } else if (numValue < 0) cellStyle.setFillForegroundColor(lightRedStyle.getFillForegroundColorColor());
                                else if (numValue == 0) cellStyle.setFillForegroundColor(lightYellowStyle.getFillForegroundColorColor());
                            } else if (columnName.equals("ROE TTM")) {
                                double roeAvg = getROEAvgForRow(tblModelRow);
                                 if (numValue > 0 && roeAvg != 0) { 
                                    double ratio = numValue / roeAvg; 
                                    if (ratio > 1.5) cellStyle.setFillForegroundColor(darkGreenStyle.getFillForegroundColorColor());
                                    else if (ratio > 1.25) cellStyle.setFillForegroundColor(mediumGreenStyle.getFillForegroundColorColor());
                                    else if (ratio > 1.0) cellStyle.setFillForegroundColor(lightGreenStyle.getFillForegroundColorColor());
                                    else if (ratio > 0.75) cellStyle.setFillForegroundColor(lightPinkStyle.getFillForegroundColorColor());
                                    else if (ratio > 0.5) cellStyle.setFillForegroundColor(mediumPinkStyle.getFillForegroundColorColor());
                                    else cellStyle.setFillForegroundColor(darkPinkStyleExcel.getFillForegroundColorColor());
                                 } else if (numValue < 0) cellStyle.setFillForegroundColor(lightRedStyle.getFillForegroundColorColor());
                                 else if (numValue == 0 && roeAvg == 0) cellStyle.setFillForegroundColor(lightYellowStyle.getFillForegroundColorColor()); 
                                 else if (numValue > 0 && numValue < 0.10) cellStyle.setFillForegroundColor(lightPinkStyle.getFillForegroundColorColor()); 
                                 else if (numValue > 0.15) cellStyle.setFillForegroundColor(darkGreenStyle.getFillForegroundColorColor()); 
                                 else if (numValue > 0.10) cellStyle.setFillForegroundColor(mediumGreenStyle.getFillForegroundColorColor());
                            } else if (columnName.equals("Debt to Equity")) {
                                double deAvg = getDEAvgForRow(tblModelRow);
                                 if (deAvg > 0 && numValue >= 0) { 
                                    double ratio = numValue / deAvg; 
                                    if (ratio < 0.5) cellStyle.setFillForegroundColor(darkGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 0.75) cellStyle.setFillForegroundColor(mediumGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.0) cellStyle.setFillForegroundColor(lightGreenStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.25) cellStyle.setFillForegroundColor(lightPinkStyle.getFillForegroundColorColor());
                                    else if (ratio < 1.5) cellStyle.setFillForegroundColor(mediumPinkStyle.getFillForegroundColorColor());
                                    else cellStyle.setFillForegroundColor(darkPinkStyleExcel.getFillForegroundColorColor());
                                 } else if (numValue < 0) cellStyle.setFillForegroundColor(darkRedStyle.getFillForegroundColorColor()); 
                                 else if (numValue >= 0 && numValue < 0.5) cellStyle.setFillForegroundColor(darkGreenStyle.getFillForegroundColorColor()); 
                                 else if (numValue < 1.0) cellStyle.setFillForegroundColor(mediumGreenStyle.getFillForegroundColorColor());
                            } else if (columnName.equals("Payout Ratio")) {
                                if (numValue < 0) cellStyle.setFillForegroundColor(darkRedStyle.getFillForegroundColorColor()); 
                                else if (numValue == 0) cellStyle.setFillForegroundColor(lightYellowStyle.getFillForegroundColorColor());
                                else if (numValue > 0 && numValue <= 0.33) cellStyle.setFillForegroundColor(darkGreenStyle.getFillForegroundColorColor());    
                                else if (numValue > 0.33 && numValue <= 0.66) cellStyle.setFillForegroundColor(mediumGreenStyle.getFillForegroundColorColor()); 
                                else if (numValue > 0.66 && numValue < 1.00) cellStyle.setFillForegroundColor(lightGreenStyle.getFillForegroundColorColor());   
                                else if (numValue >= 1.00 && numValue <= 1.25) cellStyle.setFillForegroundColor(lightRedStyle.getFillForegroundColorColor());     
                                else if (numValue > 1.25 && numValue <= 1.50) cellStyle.setFillForegroundColor(mediumRedStyle.getFillForegroundColorColor());    
                                else if (numValue > 1.50) cellStyle.setFillForegroundColor(darkRedStyle.getFillForegroundColorColor());      
                            } else if (columnName.startsWith("EPS Growth")) {
                                if (numValue < 0) cellStyle.setFillForegroundColor(lightRedStyle.getFillForegroundColorColor());
                                else if (numValue < 15) cellStyle.setFillForegroundColor(lightGreenStyle.getFillForegroundColorColor());
                                else if (numValue < 30) cellStyle.setFillForegroundColor(mediumGreenStyle.getFillForegroundColorColor());
                                else cellStyle.setFillForegroundColor(darkGreenStyle.getFillForegroundColorColor());
                            }
                            else { 
                                 if (numValue < 0) cellStyle.setFillForegroundColor(lightRedStyle.getFillForegroundColorColor());
                                 else if (numValue == 0) cellStyle.setFillForegroundColor(lightYellowStyle.getFillForegroundColorColor());
                            }
                            
                            if (cellStyle.getFillForegroundColorColor() != null && 
                                (normalStyle.getFillForegroundColorColor() == null || 
                                !cellStyle.getFillForegroundColorColor().equals(normalStyle.getFillForegroundColorColor())) ) {
                                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            }
                            cell.setCellStyle(cellStyle);

                        } else if (value != null) { // Handle non-Double values (like Refresh Date, Ticker, Name)
                            cell.setCellValue(value.toString());
                            cell.setCellStyle(normalStyle);
                        } else {
                             cell.setCellStyle(normalStyle); 
                        }
                    }
                }

                for (int i = 0; i < excelHeaderRow.getLastCellNum(); i++) {
                    sheet.autoSizeColumn(i);
                }

                try (FileOutputStream outputStream = new FileOutputStream(file)) {
                    workbook.write(outputStream);
                    JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Watchlist exported successfully!", "Export Complete", JOptionPane.INFORMATION_MESSAGE);
                }

            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(watchlistTable.getTopLevelAncestor(), "Error exporting file: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }
}
