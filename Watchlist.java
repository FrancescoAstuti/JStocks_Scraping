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
import java.awt.event.ActionListener;
import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.HttpURLConnection;
import java.net.URL;
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
import java.awt.Color;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import java.io.FileOutputStream;

public class Watchlist {

    private JTable watchlistTable;
    private DefaultTableModel tableModel;
    private static final String API_KEY = API_key.getApiKey();
    private static final String COLUMN_SETTINGS_FILE = "column_settings.properties";
    private JPanel columnControlPanel;

    private String[] getDynamicColumnNames() {
        int currentYear = Calendar.getInstance().get(Calendar.YEAR);
        String[] columnNames = new String[3];
        for (int i = 0; i < 3; i++) {
            columnNames[i] = "EPS" + (currentYear + i);
        }
        return columnNames;
    }

    public void createAndShowGUI() {

        // Set Nimbus Look and Feel before creating any GUI components
        try {
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    UIManager.setLookAndFeel(info.getClassName());
                    // Customize some Nimbus defaults if desired
                    UIManager.put("control", new Color(240, 240, 240));
                    UIManager.put("info", new Color(242, 242, 189));
                    UIManager.put("nimbusBase", new Color(51, 98, 140));
                    UIManager.put("nimbusBlueGrey", new Color(169, 176, 190));
                    UIManager.put("nimbusSelectionBackground", new Color(104, 93, 156));
                    break;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            // Fall back to system look and feel if Nimbus isn't available
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }

        JFrame frame = new JFrame("Watchlist");
        frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
        frame.setSize(1000, 600);

        JPanel mainPanel = new JPanel(new BorderLayout());

        String[] dynamicColumnNames = getDynamicColumnNames();
        String[] peForwardColumnNames = getPEForwardColumnNames();

        tableModel = new DefaultTableModel(new Object[]{
                "Name", "Ticker", "Price", "PE TTM", "PB TTM", "Div. yield %",
                "Payout Ratio", "Graham Number", "PB Avg", "PE Avg",
                "EPS TTM", "ROE TTM", "A-Score",
                dynamicColumnNames[0], dynamicColumnNames[1], dynamicColumnNames[2],
                "Debt to Equity", "EPS Growth 1", "Current Ratio", "Quick Ratio",
                "EPS Growth 2", "EPS Growth 3", "DE Avg", "Industry", "ROE Avg", "P/FCF", "PFCF Avg",
                peForwardColumnNames[0], peForwardColumnNames[1], peForwardColumnNames[2], "Volatility"}, 0) {

            @Override
            public Class<?> getColumnClass(int columnIndex) {
                switch (columnIndex) {
                    case 2:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                    case 9:
                    case 10:
                    case 11:
                    case 12:
                    case 13:
                    case 14:
                    case 15:
                    case 16:
                    case 17:
                    case 18:
                    case 19:
                    case 20:
                    case 21:
                    case 22:
                    case 23:
                    case 24:
                    case 25:
                    case 26:
                    case 27:
                    case 28:
                    case 29:
                    case 30:

                        return Double.class;
                    default:
                        return String.class;
                }
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

        watchlistTable.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 1) {  // Single click
                    int row = watchlistTable.rowAtPoint(e.getPoint());
                    int col = watchlistTable.columnAtPoint(e.getPoint());

                    if (row >= 0 && col == 1) {  // Check if click is in the Ticker column
                        String ticker = (String) watchlistTable.getValueAt(row, col);
                        // Get the company name from the first column (index 0)
                        String companyName = (String) watchlistTable.getValueAt(row, 0);
                        CompanyOverview.showCompanyOverview(ticker, companyName);
                    }
                }
            }
        });

        // Enhanced ScrollPane with explicit horizontal scroll bar policy
        JScrollPane scrollPane = new JScrollPane(
                watchlistTable,
                JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
                JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED
        );
        // Make sure the table doesn't automatically resize columns to fit the viewport
        watchlistTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        mainPanel.add(scrollPane, BorderLayout.CENTER);

        setupColumnControlPanel();
        mainPanel.add(columnControlPanel, BorderLayout.WEST);

        JPanel buttonPanel = new JPanel();
        setupButtonPanel(buttonPanel);
        mainPanel.add(buttonPanel, BorderLayout.SOUTH);

        frame.add(mainPanel);
        frame.setLocationRelativeTo(null);

        // Handle window closing
        frame.addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosing(java.awt.event.WindowEvent windowEvent) {
                int confirm = JOptionPane.showOptionDialog(
                        frame,
                        "Are you sure you want to close this window?",
                        "Close Confirmation",
                        JOptionPane.YES_NO_OPTION,
                        JOptionPane.QUESTION_MESSAGE,
                        null, null, null);
                if (confirm == JOptionPane.YES_OPTION) {
                    saveColumnSettings();
                    saveWatchlist();
                    frame.dispose();
                }
            }
        });

        // After creating the UI, load settings asynchronously
        SwingUtilities.invokeLater(() -> {
            loadColumnSettings();
            loadWatchlist();
            loadSortOrder(sorter);
        });

        frame.setVisible(true);
    }

    private void setupTableSorter(TableRowSorter<DefaultTableModel> sorter) {
        // Set comparators for numeric columns if needed
        for (int i = 2; i <= 29; i++) {
            final int column = i;
            sorter.setComparator(column, Comparator.comparingDouble(o -> (Double) o));
        }

        // Use RowSorterListener instead of propertyChangeListener
        sorter.addRowSorterListener(new RowSorterListener() {
            @Override
            public void sorterChanged(RowSorterEvent e) {
                if (e.getType() == RowSorterEvent.Type.SORT_ORDER_CHANGED) {
                    saveSortOrder(sorter.getSortKeys());
                }
            }
        });
    }

    private void setupColumnControlPanel() {
        columnControlPanel = new JPanel();
        columnControlPanel.setLayout(new BoxLayout(columnControlPanel, BoxLayout.Y_AXIS));

        TableColumnModel columnModel = watchlistTable.getColumnModel();
        List<JCheckBox> checkBoxes = new ArrayList<>();

        // Create checkboxes for each column and add them to the list
        for (int i = 0; i < columnModel.getColumnCount(); i++) {
            String columnName = tableModel.getColumnName(i);
            JCheckBox checkBox = new JCheckBox(columnName, true);
            checkBox.addActionListener(e -> toggleColumnVisibility(columnName, checkBox.isSelected()));
            checkBoxes.add(checkBox);
        }

        // Sort checkboxes alphabetically by their text
        checkBoxes.sort(Comparator.comparing(JCheckBox::getText));

        // Add sorted checkboxes to the panel
        for (JCheckBox checkBox : checkBoxes) {
            columnControlPanel.add(checkBox);
        }
    }

    private void setupButtonPanel(JPanel buttonPanel) {
        JButton addButton = new JButton("Add Stock");
        JButton deleteButton = new JButton("Delete Stock");
        JButton refreshButton = new JButton("Refresh");
        JButton exportButton = new JButton("Export XLSX");
        JButton analyticsButton = new JButton("Analytics");
        JButton dataButton = new JButton("Data");

        addButton.addActionListener(e -> addStock());
        deleteButton.addActionListener(e -> deleteStock());
        refreshButton.addActionListener(e -> refreshWatchlist());
        exportButton.addActionListener(e -> exportToExcel());
        analyticsButton.addActionListener(e -> Analytics.analyzeWatchlist());
        dataButton.addActionListener(e -> Download_Data.downloadTickerData(watchlistTable, tableModel));

        buttonPanel.add(addButton);
        buttonPanel.add(deleteButton);
        buttonPanel.add(refreshButton);
        buttonPanel.add(exportButton);
        buttonPanel.add(analyticsButton);
        buttonPanel.add(dataButton);
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
            try (FileReader reader = new FileReader(file)) {
                Scanner scanner = new Scanner(reader);
                String json = scanner.useDelimiter("\\Z").next();
                scanner.close();

                System.out.println("Loading watchlist data...");
                JSONArray jsonArray = new JSONArray(json);
                System.out.println("Found " + jsonArray.length() + " stocks");

                while (tableModel.getRowCount() > 0) {
                    tableModel.removeRow(0);
                }

                for (int i = 0; i < jsonArray.length(); i++) {
                    JSONObject jsonObject = jsonArray.getJSONObject(i);

                    Object[] rowData = new Object[]{
                            jsonObject.optString("name", ""),
                            jsonObject.optString("ticker", "").toUpperCase(),
                            jsonObject.optDouble("price", 0.0),
                            jsonObject.optDouble("peTtm", 0.0),
                            jsonObject.optDouble("pbTtm", 0.0),
                            jsonObject.optDouble("dividendYield", 0.0),
                            jsonObject.optDouble("payoutRatio", 0.0),
                            jsonObject.optDouble("grahamNumber", 0.0),
                            jsonObject.optDouble("pbAvg", 0.0),
                            jsonObject.optDouble("peAvg", 0.0),
                            jsonObject.optDouble("epsTtm", 0.0),
                            jsonObject.optDouble("roeTtm", 0.0),
                            jsonObject.optDouble("aScore", 0.0),
                            jsonObject.optDouble("epsCurrentYear", 0.0),
                            jsonObject.optDouble("epsNextYear", 0.0),
                            jsonObject.optDouble("epsYear3", 0.0),
                            jsonObject.optDouble("debtToEquity", 0.0),
                            jsonObject.optDouble("epsGrowth1", 0.0),
                            jsonObject.optDouble("currentRatio", 0.0),
                            jsonObject.optDouble("quickRatio", 0.0),
                            jsonObject.optDouble("epsGrowth2", 0.0),
                            jsonObject.optDouble("epsGrowth3", 0.0),
                            jsonObject.optDouble("deAvg", 0.0),
                            jsonObject.optString("industry", "N/A"),
                            jsonObject.optDouble("roeAvg", 0.0),
                            jsonObject.optDouble("priceToFCF_TTM", 0.0),
                            jsonObject.optDouble("priceToFCF_Avg", 0.0),
                            jsonObject.optDouble("peForward1", 0.0),
                            jsonObject.optDouble("peForward2", 0.0),
                            jsonObject.optDouble("peForward3", 0.0),
                            jsonObject.optDouble("volatility", 0.0),};
                    tableModel.addRow(rowData);
                    System.out.println("Added stock: " + jsonObject.optString("ticker", "")
                            + " with price: " + jsonObject.optDouble("price", 0.0));
                }
                System.out.println("Watchlist loading completed");
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(null,
                        "Error loading watchlist: " + e.getMessage(),
                        "Error",
                        JOptionPane.ERROR_MESSAGE);
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
            jsonObject.put("name", tableModel.getValueAt(i, 0));
            jsonObject.put("ticker", tableModel.getValueAt(i, 1));
            jsonObject.put("price", tableModel.getValueAt(i, 2));
            jsonObject.put("peTtm", tableModel.getValueAt(i, 3));
            jsonObject.put("pbTtm", tableModel.getValueAt(i, 4));
            jsonObject.put("dividendYield", tableModel.getValueAt(i, 5));
            jsonObject.put("payoutRatio", tableModel.getValueAt(i, 6));
            jsonObject.put("grahamNumber", tableModel.getValueAt(i, 7));
            jsonObject.put("pbAvg", tableModel.getValueAt(i, 8));
            jsonObject.put("peAvg", tableModel.getValueAt(i, 9));
            jsonObject.put("epsTtm", tableModel.getValueAt(i, 10));
            jsonObject.put("roeTtm", tableModel.getValueAt(i, 11));
            jsonObject.put("aScore", tableModel.getValueAt(i, 12));
            jsonObject.put("epsCurrentYear", tableModel.getValueAt(i, 13));
            jsonObject.put("epsNextYear", tableModel.getValueAt(i, 14));
            jsonObject.put("epsYear3", tableModel.getValueAt(i, 15));
            jsonObject.put("debtToEquity", tableModel.getValueAt(i, 16));
            jsonObject.put("epsGrowth1", tableModel.getValueAt(i, 17));
            jsonObject.put("currentRatio", tableModel.getValueAt(i, 18));
            jsonObject.put("quickRatio", tableModel.getValueAt(i, 19));
            jsonObject.put("epsGrowth2", tableModel.getValueAt(i, 20));
            jsonObject.put("epsGrowth3", tableModel.getValueAt(i, 21));
            jsonObject.put("deAvg", tableModel.getValueAt(i, 22));
            jsonObject.put("industry", tableModel.getValueAt(i, 23));
            jsonObject.put("roeAvg", tableModel.getValueAt(i, 24));
            jsonObject.put("priceToFCF_TTM", tableModel.getValueAt(i, 25));
            jsonObject.put("priceToFCF_Avg", tableModel.getValueAt(i, 26));
            jsonObject.put("peForward1", tableModel.getValueAt(i, 27));
            jsonObject.put("peForward2", tableModel.getValueAt(i, 28));
            jsonObject.put("peForward3", tableModel.getValueAt(i, 29));
            jsonObject.put("volatility", tableModel.getValueAt(i, 30));
            jsonArray.put(jsonObject);
        }

        try (FileWriter fileWriter = new FileWriter("watchlist.json")) {
            fileWriter.write(jsonArray.toString());
            fileWriter.flush();
            System.out.println("Watchlist saved successfully");
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,
                    "Error saving watchlist: " + e.getMessage(),
                    "Error",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void saveColumnSettings() {
        Properties properties = new Properties();
        TableColumnModel columnModel = watchlistTable.getColumnModel();

        for (int i = 0; i < columnModel.getColumnCount(); i++) {
            TableColumn column = columnModel.getColumn(i);
            String columnName = column.getHeaderValue().toString();
            properties.setProperty("column" + i + ".name", columnName);
            properties.setProperty("column" + i + ".index", String.valueOf(columnModel.getColumnIndex(columnName)));
            properties.setProperty("column" + i + ".visible", String.valueOf(column.getMaxWidth() != 0));
        }

        try (FileOutputStream out = new FileOutputStream(COLUMN_SETTINGS_FILE)) {
            properties.store(out, "Column Order and Visibility Settings");
            System.out.println("Column settings saved successfully");
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,
                    "Error saving column settings: " + e.getMessage(),
                    "Error",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void loadColumnSettings() {
        Properties properties = new Properties();
        File settingsFile = new File(COLUMN_SETTINGS_FILE);

        if (settingsFile.exists()) {
            try (FileInputStream in = new FileInputStream(settingsFile)) {
                properties.load(in);

                TableColumnModel columnModel = watchlistTable.getColumnModel();
                for (int i = 0; i < columnModel.getColumnCount(); i++) {
                    String columnName = properties.getProperty("column" + i + ".name");
                    if (columnName != null) {
                        int columnIndex = Integer.parseInt(properties.getProperty("column" + i + ".index"));
                        boolean isVisible = Boolean.parseBoolean(properties.getProperty("column" + i + ".visible"));

                        TableColumn column = columnModel.getColumn(columnModel.getColumnIndex(columnName));
                        columnModel.moveColumn(columnModel.getColumnIndex(columnName), columnIndex);
                        toggleColumnVisibility(columnName, isVisible);

                        for (Component comp : columnControlPanel.getComponents()) {
                            if (comp instanceof JCheckBox) {
                                JCheckBox checkBox = (JCheckBox) comp;
                                if (checkBox.getText().equals(columnName)) {
                                    checkBox.setSelected(isVisible);
                                    break;
                                }
                            }
                        }
                    }
                }
                System.out.println("Column settings loaded successfully");
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(null,
                        "Error loading column settings: " + e.getMessage(),
                        "Error",
                        JOptionPane.ERROR_MESSAGE);
            }
        }
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
            System.out.println("Sort order saved successfully");
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,
                    "Error saving sort order: " + e.getMessage(),
                    "Error",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void loadSortOrder(TableRowSorter<DefaultTableModel> sorter) {
        Properties properties = new Properties();
        File file = new File("sort_order.properties");

        if (file.exists()) {
            try (FileInputStream in = new FileInputStream(file)) {
                properties.load(in);

                List<RowSorter.SortKey> sortKeys = new ArrayList<>();
                for (int i = 0;; i++) {
                    String columnStr = properties.getProperty("sortKey" + i + ".column");
                    String orderStr = properties.getProperty("sortKey" + i + ".order");
                    if (columnStr == null || orderStr == null) {
                        break;
                    }

                    int column = Integer.parseInt(columnStr);
                    SortOrder order = SortOrder.valueOf(orderStr);
                    sortKeys.add(new RowSorter.SortKey(column, order));
                }

                sorter.setSortKeys(sortKeys);
                System.out.println("Sort order loaded successfully");
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(null,
                        "Error loading sort order: " + e.getMessage(),
                        "Error",
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void addStock() {
        String ticker = JOptionPane.showInputDialog("Enter Stock Ticker:");
        if (ticker != null && !ticker.trim().isEmpty()) {
            ticker = ticker.toUpperCase();

            // Check if ticker already exists in the watchlist
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                if (ticker.equals((String) tableModel.getValueAt(i, 1))) {
                    JOptionPane.showMessageDialog(null,
                            "Stock '" + ticker + "' is already in the watchlist.",
                            "Duplicate Stock",
                            JOptionPane.WARNING_MESSAGE);
                    return;
                }
            }

            try {
                JSONObject stockData = fetchStockData(ticker);

                if (stockData == null) {
                    JOptionPane.showMessageDialog(null,
                            "Symbol '" + ticker + "' not found.",
                            "Symbol Not Found",
                            JOptionPane.WARNING_MESSAGE);
                    return;
                } else {
                    System.out.println("Stock data fetched for ticker: " + ticker);
                }

                JSONObject ratios = fetchStockRatios(ticker);
                if (ratios == null) {
                    JOptionPane.showMessageDialog(null,
                            "Financial data not available for '" + ticker + "'.",
                            "Data Not Available",
                            JOptionPane.WARNING_MESSAGE);
                    return;
                } else {
                    System.out.println("Ratios data fetched for ticker: " + ticker);
                }

                // Assuming stockData and ratios contain the necessary fields
                String industry = CompanyOverview.fetchIndustry(ticker);
                String name = stockData.getString("name");
                double price = stockData.getDouble("price");
                double peTtm = ratios.optDouble("peRatioTTM", 0.0);
                double pbTtm = ratios.optDouble("pbRatioTTM", 0.0);
                double dividendYieldTTM = KeyMetricsTTM.getDividendYieldTTM(ticker);
                double payoutRatio = ratios.optDouble("payoutRatioTTM", 0.0);
                double pbAvg = fetchAveragePB(ticker);
                double peAvg = fetchAveragePE(ticker);
                double epsTtm = peTtm != 0 ? round((1 / peTtm) * price, 2) : 0.0;
                double grahamNumber = calculateGrahamNumber(price, peAvg, pbAvg, epsTtm, pbTtm);
                double roeTtm = round(ratios.optDouble("roeTTM", 0.0), 2);
                double epsCurrentYear = ratios.optDouble("epsCurrentYear", 0.0);
                double epsNextYear = ratios.optDouble("epsNextYear", 0.0);
                double epsYear3 = ratios.optDouble("epsYear3", 0.0);
                double debtToEquity = KeyMetricsTTM.getDebtToEquityTTM(ticker);
                double epsGrowth1 = calculateEpsGrowth1(epsCurrentYear, epsTtm);
                double currentRatio = ratios.optDouble("currentRatioTTM", 0.0);
                double quickRatio = ratios.optDouble("quickRatioTTM", 0.0);
                double epsGrowth2 = calculateEpsGrowth2(epsCurrentYear, epsNextYear);
                double epsGrowth3 = calculateEpsGrowth2(epsNextYear, epsYear3);
                double deAvg = Ratios.fetchDebtToEquityAverage(ticker);
                double roeAvg = 0;
                double aScore = 0;

                Object[] rowData = new Object[]{
                        name, ticker, price, peTtm, pbTtm, dividendYieldTTM, payoutRatio, grahamNumber, pbAvg, peAvg, epsTtm, roeTtm, aScore,
                        epsCurrentYear, epsNextYear, epsYear3, debtToEquity, epsGrowth1, currentRatio, quickRatio, epsGrowth2, epsGrowth3, deAvg, industry,};

                tableModel.addRow(rowData);
                System.out.println("Added stock: " + ticker + " deAvg" + deAvg);

                // Save the updated watchlist
                saveWatchlist();

            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(null,
                        "Error processing '" + ticker + "': " + e.getMessage(),
                        "Error",
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private double calculateEpsGrowth1(double epsCurrentYear, double epsTtm) {
        if (epsTtm == 0) {
            return 0;
        } else if (epsTtm < 0) {
            double growthRate1 = 100 * (epsCurrentYear - epsTtm) / (-epsTtm);

            return round(growthRate1, 2);
        } else {
            double growthRate1 = 100 * (epsCurrentYear - epsTtm) / epsTtm;

            return round(growthRate1, 2);
        }
    }

    private double calculateEpsGrowth2(double epsCurrentYear, double epsNextYear) {
        if (epsCurrentYear == 0) {
            return 0;
        }
        double growthRate2 = 100 * (epsNextYear - epsCurrentYear) / epsCurrentYear;

        return round(growthRate2, 2);
    }

    private double calculateEpsGrowth3(double epsYear3, double epsNextYear) {
        if (epsNextYear == 0) {
            return 0;
        }
        double growthRate3 = 100 * (epsYear3 - epsNextYear) / epsNextYear;

        return round(growthRate3, 2);
    }

    private void refreshWatchlist() {
        System.out.println("Starting watchlist refresh...");

        // Create progress bar panel
        JPanel progressPanel = new JPanel(new BorderLayout());
        JProgressBar progressBar = new JProgressBar(0, tableModel.getRowCount());
        JLabel statusLabel = new JLabel("Refreshing watchlist...");
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
                        SwingUtilities.invokeLater(() -> {
                            progressBar.setValue(progress);
                            statusLabel.setText("Refreshing: " + ticker + " (" + progress + "/" + rowCount + ")");
                        });

                        JSONObject stockData = fetchStockData(ticker);
                        JSONObject ratios = fetchStockRatios(ticker);
                        JSONObject epsEstimates = Estimates.fetchEpsEstimates(ticker);

                        if (stockData != null) {
                            final double price = convertGbxToGbp(round(stockData.getDouble("price"), 2), ticker);
                            double peTtm = round(ratios.optDouble("peRatioTTM", 0.0), 2);
                            double pbTtm = round(ratios.optDouble("pbRatioTTM", 0.0), 2);
                            double epsTtm = peTtm != 0 ? round((1 / peTtm) * price, 2) : 0.0;
                            double roeTtm = round(ratios.optDouble("roeTTM", 0.0), 2);
                            double dividendYieldTTM = KeyMetricsTTM.getDividendYieldTTM(ticker);
                            double payoutRatio = round(ratios.optDouble("payoutRatioTTM", 0.0), 2);
                            double debtToEquity = KeyMetricsTTM.getDebtToEquityTTM(ticker);
                            double epsCurrentYear = epsEstimates != null
                                    ? round(epsEstimates.optDouble("eps0", 0.0), 2)
                                    : 0.0;
                            double epsNextYear = epsEstimates != null
                                    ? round(epsEstimates.optDouble("eps1", 0.0), 2)
                                    : 0.0;
                            double epsYear3 = epsEstimates != null
                                    ? round(epsEstimates.optDouble("eps2", 0.0), 2)
                                    : 0.0;
                            double currentRatio = ratios != null
                                    ? round(ratios.optDouble("currentRatioTTM", 0.0), 2)
                                    : 0.0;
                            double quickRatio = ratios != null
                                    ? round(ratios.optDouble("quickRatioTTM", 0.0), 2)
                                    : 0.0;

                            double epsGrowth1 = calculateEpsGrowth1(epsCurrentYear, epsTtm);
                            double epsGrowth2 = calculateEpsGrowth2(epsCurrentYear, epsNextYear);
                            double epsGrowth3 = calculateEpsGrowth3(epsYear3, epsNextYear);
                            //* double pbAvg = fetchAveragePB(ticker);
                            double peForward1 = calculatePEForward(price, epsCurrentYear);
                            double peForward2 = calculatePEForward(price, epsNextYear);
                            double peForward3 = calculatePEForward(price, epsYear3);
                            //* double peAvg = fetchAveragePE(ticker);
                            double priceToFCF_TTM = TTM_Ratios.getPriceToFreeCashFlowRatioTTM(ticker);
                            //* double PriceToFCF_Avg = Ratios.fetchPriceToFreeCashFlowAverage(ticker);
                            //* double roeAvg = fetchAverageROE(ticker);
                            //* double grahamNumber = calculateGrahamNumber(price, peAvg, pbAvg, epsTtm, pbTtm);
                            //* double deAvg = Ratios.fetchDebtToEquityAverage(ticker);
                            String industry = CompanyOverview.fetchIndustry(ticker);
                            //* double aScore = calculateAScore(pbAvg, pbTtm, peAvg, peTtm, payoutRatio, debtToEquity, roeTtm, roeAvg, dividendYieldTTM, deAvg, epsGrowth1, epsGrowth2, epsGrowth3, currentRatio, quickRatio, grahamNumber, price, priceToFCF_TTM, PriceToFCF_Avg);
                            double volatilityScore = Analytics.calculateVolatilityScore(ticker);
                            //*System.out.printf("Ticker: %s, DebtToEquity: %s, A-Score: %f%n", ticker, debtToEquity, aScore);

                            SwingUtilities.invokeLater(() -> {
                                tableModel.setValueAt(price, modelRow, 2);
                                tableModel.setValueAt(peTtm, modelRow, 3);
                                tableModel.setValueAt(pbTtm, modelRow, 4);
                                tableModel.setValueAt(dividendYieldTTM, modelRow, 5);
                                tableModel.setValueAt(payoutRatio, modelRow, 6);
                                //* tableModel.setValueAt(grahamNumber, modelRow, 7);
                                //*tableModel.setValueAt(pbAvg, modelRow, 8); // PB Avg
                                //* tableModel.setValueAt(peAvg, modelRow, 9); // PE Avg
                                tableModel.setValueAt(epsTtm, modelRow, 10);
                                tableModel.setValueAt(roeTtm, modelRow, 11);
                                //* tableModel.setValueAt(aScore, modelRow, 12);
                                tableModel.setValueAt(epsCurrentYear, modelRow, 13);
                                tableModel.setValueAt(epsNextYear, modelRow, 14);
                                tableModel.setValueAt(epsYear3, modelRow, 15);
                                tableModel.setValueAt(debtToEquity, modelRow, 16);
                                tableModel.setValueAt(epsGrowth1, modelRow, 17);
                                tableModel.setValueAt(currentRatio, modelRow, 18); // Index of the new "Current Ratio" column
                                tableModel.setValueAt(quickRatio, modelRow, 19); // Index of the new "Quick Ratio" column
                                tableModel.setValueAt(epsGrowth2, modelRow, 20);
                                tableModel.setValueAt(epsGrowth3, modelRow, 21);
                                //*tableModel.setValueAt(deAvg, modelRow, 22);
                                tableModel.setValueAt(industry, modelRow, 23);
                                //* tableModel.setValueAt(roeAvg, modelRow, 24); // ROE Avg
                                tableModel.setValueAt(priceToFCF_TTM, modelRow, 25);
                                //*tableModel.setValueAt(PriceToFCF_Avg, modelRow, 26);
                                tableModel.setValueAt(peForward1, modelRow, 27);
                                tableModel.setValueAt(peForward2, modelRow, 28);
                                tableModel.setValueAt(peForward3, modelRow, 29);
                                tableModel.setValueAt(volatilityScore, modelRow, 30);

                            });

                            System.out.println("Refreshed stock data: " + ticker);
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                        System.err.println("Error refreshing " + ticker + ": " + e.getMessage());
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

                saveWatchlist();
                System.out.println("Watchlist refresh completed");

                JOptionPane.showMessageDialog(
                        watchlistTable,
                        "Refresh completed!",
                        "Watchlist",
                        JOptionPane.INFORMATION_MESSAGE);

            }
        };

        worker.execute();
    }

    private void deleteStock() {
        int selectedRow = watchlistTable.getSelectedRow();
        if (selectedRow != -1) {
            int modelRow = watchlistTable.convertRowIndexToModel(selectedRow);
            String ticker = (String) tableModel.getValueAt(modelRow, 1);
            tableModel.removeRow(modelRow);
            saveWatchlist();
            System.out.println("Stock deleted successfully: " + ticker);
        }
    }

    private double round(double value, int places) {
        if (places < 0) {
            throw new IllegalArgumentException();
        }
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
                Scanner scanner = new Scanner(url.openStream());
                String response = scanner.useDelimiter("\\Z").next();
                scanner.close();

                JSONArray data = new JSONArray(response);
                if (data.length() > 0) {
                    return data.getJSONObject(0);
                } else {
                    return null;
                }
            }
            return null;
        } catch (Exception e) {
            return null;
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }
    }

    private JSONObject fetchStockRatios(String ticker) {
        // Create a combined JSONObject to store results from both endpoints
        JSONObject combinedRatios = new JSONObject();

        // First endpoint (key-metrics-ttm)
        String urlMetrics = String.format("https://financialmodelingprep.com/api/v3/key-metrics-ttm/%s?apikey=%s", ticker, API_KEY);
        // Second endpoint (ratios-ttm)
        String urlRatios = String.format("https://financialmodelingprep.com/api/v3/ratios-ttm/%s?apikey=%s", ticker, API_KEY);

        try {
            // Fetch data from first endpoint (key-metrics-ttm)
            URL url1 = new URL(urlMetrics);
            HttpURLConnection conn1 = (HttpURLConnection) url1.openConnection();
            conn1.setRequestMethod("GET");

            if (conn1.getResponseCode() == 200) {
                Scanner scanner1 = new Scanner(url1.openStream());
                String response1 = scanner1.useDelimiter("\\Z").next();
                scanner1.close();

                JSONArray data1 = new JSONArray(response1);
                if (data1.length() > 0) {
                    // Copy all properties from first endpoint
                    JSONObject metricsData = data1.getJSONObject(0);
                    for (String key : metricsData.keySet()) {
                        combinedRatios.put(key, metricsData.get(key));
                    }
                }
            }
            conn1.disconnect();

            // Fetch data from second endpoint (ratios-ttm)
            URL url2 = new URL(urlRatios);
            HttpURLConnection conn2 = (HttpURLConnection) url2.openConnection();
            conn2.setRequestMethod("GET");

            if (conn2.getResponseCode() == 200) {
                Scanner scanner2 = new Scanner(url2.openStream());
                String response2 = scanner2.useDelimiter("\\Z").next();
                scanner2.close();

                JSONArray data2 = new JSONArray(response2);
                if (data2.length() > 0) {
                    // Copy all properties from second endpoint
                    JSONObject ratiosData = data2.getJSONObject(0);
                    for (String key : ratiosData.keySet()) {
                        // Only overwrite if the value doesn't exist or is 0
                        if (!combinedRatios.has(key) || combinedRatios.getDouble(key) == 0) {
                            combinedRatios.put(key, ratiosData.get(key));
                        }
                    }
                }
            }
            conn2.disconnect();

            return combinedRatios;

        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    private double fetchAveragePB(String ticker) {
        String urlString = String.format("https://financialmodelingprep.com/api/v3/key-metrics/%s?period=annual&limit=20&apikey=%s", ticker, API_KEY);
        HttpURLConnection connection = null;
        double sum = 0;
        int count = 0;

        try {
            URL url = new URL(urlString);
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");

            int responseCode = connection.getResponseCode();
            if (responseCode == 200) {
                Scanner scanner = new Scanner(url.openStream());
                String response = scanner.useDelimiter("\\Z").next();
                scanner.close();

                JSONArray data = new JSONArray(response);
                for (int i = 0; i < data.length(); i++) {
                    JSONObject metrics = data.getJSONObject(i);
                    double pbRatio = metrics.optDouble("pbRatio", 0.0);
                    if (pbRatio != 0.0) {

                        if (pbRatio > 0) {
                            pbRatio = Math.min(pbRatio, 10.0);
                        } else {
                            pbRatio = Math.max(pbRatio, -10.0);
                        }

                        sum += pbRatio;
                        count++;
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }

        return count > 0 ? round(sum / count, 2) : 0.0;
    }

    private double fetchAveragePE(String ticker) {
        String urlString = String.format("https://financialmodelingprep.com/api/v3/key-metrics/%s?period=annual&limit=20&apikey=%s", ticker, API_KEY);
        HttpURLConnection connection = null;
        double sum = 0;
        int count = 0;

        try {
            URL url = new URL(urlString);
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");

            int responseCode = connection.getResponseCode();
            if (responseCode == 200) {
                Scanner scanner = new Scanner(url.openStream());
                String response = scanner.useDelimiter("\\Z").next();
                scanner.close();

                JSONArray data = new JSONArray(response);
                for (int i = 0; i < data.length(); i++) {
                    JSONObject metrics = data.getJSONObject(i);
                    double peRatio = metrics.optDouble("peRatio", 0.0);
                    if (peRatio != 0.0) {  // Include all non-zero PE ratios
                        // Cap PE ratios: positive at 30, negative at -30
                        if (peRatio > 0) {
                            peRatio = Math.min(peRatio, 30.0);
                        } else {
                            peRatio = Math.max(peRatio, -30.0);
                        }
                        sum += peRatio;
                        count++;
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }

        return count > 0 ? round(sum / count, 2) : 0.0;
    }

    private double fetchAverageROE(String ticker) {
        String urlString = String.format("https://financialmodelingprep.com/api/v3/key-metrics/%s?period=annual&limit=20&apikey=%s", ticker, API_KEY);
        HttpURLConnection connection = null;
        double sum = 0;
        int count = 0;

        try {
            URL url = new URL(urlString);
            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");

            int responseCode = connection.getResponseCode();
            if (responseCode == 200) {
                Scanner scanner = new Scanner(url.openStream());
                String response = scanner.useDelimiter("\\Z").next();
                scanner.close();

                JSONArray data = new JSONArray(response);
                for (int i = 0; i < data.length(); i++) {
                    JSONObject metrics = data.getJSONObject(i);
                    double roe = metrics.optDouble("roe", 0.0);
                    if (roe != 0.0) {
                        sum += roe;
                        count++;
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }

        return count > 0 ? round(sum / count, 2) : 0.0;
    }

    private double calculateGrahamNumber(double price, double peAvg, double pbAvg, double epsTtm, double pbTtm) {
        if (peAvg <= 0 || pbAvg <= 0 || price <= 0 || pbTtm <= 0 || epsTtm <= 0) {
            return 0.0;
        }

        // Calculate Graham Number using both average and current values
        // The formula: sqrt(peAvg * pbAvg * epsTtm * (1/pbTtm) * price)
        double grahamNumber = Math.sqrt(peAvg * pbAvg * epsTtm * (1 / pbTtm) * price);

        return round(grahamNumber, 2);
    }

    private double convertGbxToGbp(double price, String ticker) {
        // Check if the ticker is from London Stock Exchange (usually ends with .L)
        if (ticker != null && ticker.endsWith(".L")) {
            // Convert from GBX to GBP (divide by 100)
            return round(price / 100, 2);
        }
        return price;
    }

    private double calculateAScore(double pbAvg, double pbTtm, double peAvg, double peTtm, double payoutRatio, double debtToEquity,
                                   double roe, double roeAvg, double dividendYieldTTM, double deAvg, double epsGrowth1, double epsGrowth2, double epsGrowth3,
                                   double currentRatio, double quickRatio, double grahamNumberTerm, double price, double priceToFCF_TTM, double priceToFCF_AVG) {
        double peRatioTerm = 0;
        double pbRatioTerm = 0;
        double payoutRatioTerm = 0;
        double dividendYieldTerm = 0;
        double debtToEquityTerm = 0;
        double deAvgTerm = 0;
        double roeTerm = 0;
        double roeAvgTerm = 0;
        double epsGrowth1Term = 0;
        double epsGrowth2Term = 0;
        double epsGrowth3Term = 0;
        double currentRatioTerm = 0;
        double quickRatioTerm = 0;
        double pfcfTerm = 0;

// Conditions for debt to Equity
        if (debtToEquity == 0 || deAvg == 0.0) {
            deAvgTerm = 0;
        } else {
            double ratio = deAvg / debtToEquity;
            if (ratio < 1) {
                deAvgTerm = 0;
            } else if (ratio >= 1 && ratio < 1.5) {
                deAvgTerm = 1;
            } else if (ratio >= 1.5) {
                deAvgTerm = 2;
            }
        }

        // Conditions for peRatioTerm
        if (peTtm <= 0) {
            peRatioTerm = -1;
        } else if (peAvg / peTtm < 1) {
            peRatioTerm = 0;
        } else if (peAvg / peTtm >= 1 && peAvg / peTtm < 1.5) {
            peRatioTerm = 1;
        } else if (peAvg / peTtm >= 1.5) {
            peRatioTerm = 2;
        }

        // Conditions for pbRatioTerm
        if (pbTtm <= 0 || pbAvg / pbTtm < 1) {
            pbRatioTerm = -2;
        } else if (pbAvg / pbTtm >= 1 && pbAvg / pbTtm < 1.5) {
            pbRatioTerm = 1;
        } else if (pbAvg / pbTtm >= 1.5) {
            pbRatioTerm = 2;
        }

        if (priceToFCF_TTM <= 0) {
            pfcfTerm = -2;  // Penalty for negative FCF
        } else if (priceToFCF_AVG / priceToFCF_TTM < 1) {
            pfcfTerm = -1;    // Current PFCF is higher than average (potentially overvalued)
        } else if (priceToFCF_AVG / priceToFCF_TTM >= 1 && priceToFCF_AVG / priceToFCF_TTM < 1.5) {
            pfcfTerm = 1;      // Current PFCF is lower than average but not significantly
        } else if (priceToFCF_AVG / priceToFCF_TTM >= 1.5) {
            pfcfTerm = 2;      // Current PFCF is significantly lower than average (potentially undervalued)
        }

        // Conditions for dividendYieldTerm
// Conditions for dividendYieldTerm
        if (dividendYieldTTM < 3) {
            dividendYieldTerm = 0;
        } else if (dividendYieldTTM >= 3 && dividendYieldTTM < 6) {
            dividendYieldTerm = 1;
        } else if (dividendYieldTTM >= 6) {
            dividendYieldTerm = 2;
        }

        // Conditions for payoutRatioTerm
        if (payoutRatio <= 0 || payoutRatio >= 1) {
            payoutRatioTerm = 0;
        } else if (payoutRatio >= 0.5 && payoutRatio < 1) {
            payoutRatioTerm = 1;
        } else {
            payoutRatioTerm = 2;
        }

        // Conditions for debtToEquityTerm
        if (debtToEquity <= 0 || (double) debtToEquity > 1) {
            debtToEquityTerm = 0;
        } else if ((double) debtToEquity >= 0.5 && (double) debtToEquity <= 1) {
            debtToEquityTerm = 1;
        } else {
            debtToEquityTerm = 2;
        }

        // Conditions for roe
        if (roe < 0.1) {
            roeTerm = 0;
        } else if (roe >= 0.1 && roe < 0.2) {
            roeTerm = 1;
        } else if (roe >= 0.2) {
            roeTerm = 2;
        }

        // Conditions for roeAvg
        if (roe <= 0 || roeAvg <= 0) {
            roeAvgTerm = -1;
        } else if (roe / roeAvg < 1) {
            roeAvgTerm = 0;
        } else if (roe / roeAvg >= 1 && roe / roeAvg < 1.5) {
            roeAvgTerm = 1;
        } else if (roe / roeAvg >= 1.5) {
            roeAvgTerm = 2;
        }

        // Conditions for epsGrowht1
        if (epsGrowth1 <= -25) {
            epsGrowth1Term = -2;
        } else if (epsGrowth1 > -25 && epsGrowth1 <= 0) {
            epsGrowth1Term = -1;
        } else if (epsGrowth1 > 0 && epsGrowth1 < 25) {
            epsGrowth1Term = 0;
        } else if (epsGrowth1 >= 25 && epsGrowth1 < 75) {
            epsGrowth1Term = 1;
        } else if (epsGrowth1 >= 75) {
            epsGrowth1Term = 2;

        }

        // Conditions for epsGrowht2
        if (epsGrowth2 <= -25) {
            epsGrowth2Term = -2;
        } else if (epsGrowth2 > -25 && epsGrowth2 <= 0) {
            epsGrowth2Term = -1;
        } else if (epsGrowth2 > 0 && epsGrowth2 < 25) {
            epsGrowth2Term = 0;
        } else if (epsGrowth2 >= 25 && epsGrowth2 < 75) {
            epsGrowth2Term = 1;
        } else if (epsGrowth2 >= 75) {
            epsGrowth2Term = 2;
        }

        // Conditions for epsGrowht2
        if (epsGrowth3 <= -25) {
            epsGrowth3Term = -2;
        } else if (epsGrowth3 > -25 && epsGrowth3 <= 0) {
            epsGrowth3Term = -1;
        } else if (epsGrowth3 > 0 && epsGrowth3 < 25) {
            epsGrowth3Term = 0;
        } else if (epsGrowth3 >= 25 && epsGrowth3 < 75) {
            epsGrowth3Term = 1;
        } else if (epsGrowth3 >= 75) {
            epsGrowth3Term = 2;
        }

        // Conditions for Graham Number Term
        double percentDiff = (grahamNumberTerm - price) / price * 100;

        if (percentDiff <= -25) {
            grahamNumberTerm = -2;
        } else if (percentDiff > -25 && percentDiff <= 0) {
            grahamNumberTerm = -1;
        } else if (percentDiff > 0 && percentDiff <= 25) {
            grahamNumberTerm = 0;
        } else if (percentDiff > 25 && percentDiff <= 50) {
            grahamNumberTerm = 1;
        } else if (percentDiff > 50) {
            grahamNumberTerm = 2;
        }

        // Conditions for current ratio
        if (currentRatio < 1) {
            currentRatioTerm = 0;
        } else if (currentRatio >= 1 && currentRatio < 2) {
            currentRatioTerm = 1;
        } else if (currentRatio >= 2) {
            currentRatioTerm = 2;
        }

        // Conditions for quick ratio
        if (quickRatio < 1) {
            quickRatioTerm = 0;
        } else if (quickRatio >= 1 && quickRatio < 2) {
            quickRatioTerm = 1;
        } else if (quickRatio >= 2) {
            quickRatioTerm = 2;
        }

        return (payoutRatioTerm + deAvgTerm + debtToEquityTerm + dividendYieldTerm + peRatioTerm + pbRatioTerm
                + payoutRatioTerm + epsGrowth1Term + epsGrowth3Term + epsGrowth2Term + currentRatioTerm
                + quickRatioTerm + roeTerm + roeAvgTerm + grahamNumberTerm + pfcfTerm);

    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new Watchlist().createAndShowGUI());
    }

    private void exportToExcel() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save Excel File");
        fileChooser.setSelectedFile(new File("watchlist.xlsx"));

        if (fileChooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();
            if (!file.getName().endsWith(".xlsx")) {
                file = new File(file.getAbsolutePath() + ".xlsx");
            }

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Watchlist");

                // Create all styles
                XSSFCellStyle headerStyle = workbook.createCellStyle();
                headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                headerStyle.setBorderBottom(BorderStyle.THIN);

                // Create all custom color styles
                XSSFCellStyle normalStyle = workbook.createCellStyle();
                XSSFCellStyle lightRedStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 255, (byte) 235, (byte) 235});
                XSSFCellStyle lightYellowStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 255, (byte) 255, (byte) 220});
                XSSFCellStyle lightGreenStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 220, (byte) 255, (byte) 220});
                XSSFCellStyle mediumGreenStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 198, (byte) 255, (byte) 198});
                XSSFCellStyle darkGreenStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 178, (byte) 255, (byte) 178});
                XSSFCellStyle lightPinkStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 255, (byte) 230, (byte) 230});
                XSSFCellStyle mediumPinkStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 255, (byte) 200, (byte) 200});
                XSSFCellStyle darkPinkStyle = CellsFormat.createCustomColorStyle(workbook, new byte[]{(byte) 255, (byte) 180, (byte) 180});

                // Get current column order
                TableColumnModel columnModel = watchlistTable.getColumnModel();
                int columnCount = columnModel.getColumnCount();

                // Create header row
                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < columnCount; i++) {
                    TableColumn tableColumn = columnModel.getColumn(i);
                    int modelIndex = tableColumn.getModelIndex();
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(tableModel.getColumnName(modelIndex));
                    cell.setCellStyle(headerStyle);
                }

                // Create data rows
                for (int row = 0; row < watchlistTable.getRowCount(); row++) {
                    Row dataRow = sheet.createRow(row + 1);
                    int modelRow = watchlistTable.convertRowIndexToModel(row);

                    for (int viewCol = 0; viewCol < columnCount; viewCol++) {
                        TableColumn tableColumn = columnModel.getColumn(viewCol);
                        int modelCol = tableColumn.getModelIndex();
                        String columnName = tableModel.getColumnName(modelCol);
                        Object value = tableModel.getValueAt(modelRow, modelCol);
                        Cell cell = dataRow.createCell(viewCol);

                        if (value instanceof Number) {
                            double numValue = ((Number) value).doubleValue();
                            cell.setCellValue(numValue);

                            if (columnName.equals("Graham Number")) {
                                double price = getPriceForRow(modelRow);
                                if (price > 0) {
                                    double percentDiff = (numValue - price) / price * 100;

                                    if (percentDiff > 0) {
                                        if (percentDiff <= 25) {
                                            cell.setCellStyle(lightGreenStyle);
                                        } else if (percentDiff <= 50) {
                                            cell.setCellStyle(mediumGreenStyle);
                                        } else {
                                            cell.setCellStyle(darkGreenStyle);
                                        }
                                    } else {
                                        percentDiff = Math.abs(percentDiff);
                                        if (percentDiff <= 25) {
                                            cell.setCellStyle(lightPinkStyle);
                                        } else if (percentDiff <= 50) {
                                            cell.setCellStyle(mediumPinkStyle);
                                        } else {
                                            cell.setCellStyle(darkPinkStyle);
                                        }
                                    }
                                }
                            } else if (columnName.equals("PE TTM")) {
                                double peAvg = getPEAvgForRow(modelRow);
                                if (peAvg > 0 && numValue > 0) {
                                    double ratio = numValue / peAvg;

                                    if (ratio < 1) {  // PE TTM is lower than average (potentially undervalued)
                                        if (ratio >= 0.75) {
                                            cell.setCellStyle(lightGreenStyle);
                                        } else if (ratio >= 0.5) {
                                            cell.setCellStyle(mediumGreenStyle);
                                        } else {
                                            cell.setCellStyle(darkGreenStyle);
                                        }
                                    } else {  // PE TTM is higher than average (potentially overvalued)
                                        if (ratio <= 1.25) {
                                            cell.setCellStyle(lightPinkStyle);
                                        } else if (ratio <= 1.5) {
                                            cell.setCellStyle(mediumPinkStyle);
                                        } else {
                                            cell.setCellStyle(darkPinkStyle);
                                        }
                                    }
                                }
                            } else if (columnName.startsWith("PE FWD")) {
                                double peAvg = getPEAvgForRow(modelRow);
                                if (peAvg > 0 && numValue > 0) {
                                    double ratio = numValue / peAvg;

                                    if (ratio < 1) {  // Forward PE is lower than average (potentially undervalued)
                                        if (ratio >= 0.75) {
                                            cell.setCellStyle(lightGreenStyle);
                                        } else if (ratio >= 0.5) {
                                            cell.setCellStyle(mediumGreenStyle);
                                        } else {
                                            cell.setCellStyle(darkGreenStyle);
                                        }
                                    } else {  // Forward PE is higher than average (potentially overvalued)
                                        if (ratio <= 1.25) {
                                            cell.setCellStyle(lightPinkStyle);
                                        } else if (ratio <= 1.5) {
                                            cell.setCellStyle(mediumPinkStyle);
                                        } else {
                                            cell.setCellStyle(darkPinkStyle);
                                        }
                                    }
                                }
                            } else if (columnName.equals("PB TTM")) {
                                double pbAvg = getPBAvgForRow(modelRow);
                                if (pbAvg > 0 && numValue > 0) {
                                    double ratio = numValue / pbAvg;

                                    if (ratio < 1) {  // PB TTM is lower than average (potentially undervalued)
                                        if (ratio >= 0.75) {
                                            cell.setCellStyle(lightGreenStyle);
                                        } else if (ratio >= 0.5) {
                                            cell.setCellStyle(mediumGreenStyle);
                                        } else {
                                            cell.setCellStyle(darkGreenStyle);
                                        }
                                    } else {  // PB TTM is higher than average (potentially overvalued)
                                        if (ratio <= 1.25) {
                                            cell.setCellStyle(lightPinkStyle);
                                        } else if (ratio <= 1.5) {
                                            cell.setCellStyle(mediumPinkStyle);
                                        } else {
                                            cell.setCellStyle(darkPinkStyle);
                                        }
                                    }
                                }
                            } else if (columnName.equals("P/FCF")) {
                                double pfcfAvg = getPriceToFCFAvgForRow(modelRow);
                                if (pfcfAvg > 0 && numValue > 0) {
                                    double ratio = numValue / pfcfAvg;

                                    if (ratio < 1) {  // P/FCF TTM is lower than average (better valuation)
                                        if (ratio >= 0.75) {
                                            cell.setCellStyle(lightGreenStyle);
                                        } else if (ratio >= 0.5) {
                                            cell.setCellStyle(mediumGreenStyle);
                                        } else {
                                            cell.setCellStyle(darkGreenStyle);
                                        }
                                    } else {  // P/FCF TTM is higher than average (worse valuation)
                                        if (ratio <= 1.25) {
                                            cell.setCellStyle(lightPinkStyle);
                                        } else if (ratio <= 1.5) {
                                            cell.setCellStyle(mediumPinkStyle);
                                        } else {
                                            cell.setCellStyle(darkPinkStyle);
                                        }
                                    }
                                }
                            } else if (columnName.equals("ROE TTM")) {
                                double roeAvg = getROEAvgForRow(modelRow);
                                if (roeAvg > 0 && numValue > 0) {
                                    double ratio = numValue / roeAvg;

                                    if (ratio > 1) {  // ROE TTM is higher than average (better performance)
                                        if (ratio <= 1.25) {
                                            cell.setCellStyle(lightGreenStyle);
                                        } else if (ratio <= 1.5) {
                                            cell.setCellStyle(mediumGreenStyle);
                                        } else {
                                            cell.setCellStyle(darkGreenStyle);
                                        }
                                    } else {  // ROE TTM is lower than average (worse performance)
                                        if (ratio >= 0.75) {
                                            cell.setCellStyle(lightPinkStyle);
                                        } else if (ratio >= 0.5) {
                                            cell.setCellStyle(mediumPinkStyle);
                                        } else {
                                            cell.setCellStyle(darkPinkStyle);
                                        }
                                    }
                                }
                            } else if (columnName.equals("Debt to Equity")) {
                                double deAvg = getDEAvgForRow(modelRow);
                                if (deAvg > 0 && numValue > 0) {
                                    double ratio = numValue / deAvg;

                                    if (ratio < 1) {  // Debt to Equity is lower than average (better)
                                        if (ratio >= 0.75) {
                                            cell.setCellStyle(lightGreenStyle);
                                        } else if (ratio >= 0.5) {
                                            cell.setCellStyle(mediumGreenStyle);
                                        } else {
                                            cell.setCellStyle(darkGreenStyle);
                                        }
                                    } else {  // Debt to Equity is higher than average (worse)
                                        if (ratio <= 1.25) {
                                            cell.setCellStyle(lightPinkStyle);
                                        } else if (ratio <= 1.5) {
                                            cell.setCellStyle(mediumPinkStyle);
                                        } else {
                                            cell.setCellStyle(darkPinkStyle);
                                        }
                                    }
                                }
                            } else {
                                // Normal number coloring
                                if (numValue < 0) {
                                    cell.setCellStyle(lightRedStyle);
                                } else if (numValue == 0.0) {
                                    cell.setCellStyle(lightYellowStyle);
                                } else {
                                    cell.setCellStyle(normalStyle);
                                }
                            }
                        } else if (value != null) {
                            cell.setCellValue(value.toString());
                            cell.setCellStyle(normalStyle);
                        } else {
                            cell.setCellStyle(normalStyle);
                        }
                    }
                }

                // Format numbers
                DataFormat format = workbook.createDataFormat();
                for (int row = 1; row <= tableModel.getRowCount(); row++) {
                    for (int col = 0; col < columnCount; col++) {
                        Cell cell = sheet.getRow(row).getCell(col);
                        if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                            XSSFCellStyle currentStyle = (XSSFCellStyle) cell.getCellStyle();
                            XSSFCellStyle newStyle = workbook.createCellStyle();
                            newStyle.cloneStyleFrom(currentStyle);
                            newStyle.setDataFormat(format.getFormat("0.00"));
                            cell.setCellStyle(newStyle);
                        }
                    }
                }

                // Auto-size columns
                for (int i = 0; i < columnCount; i++) {
                    TableColumn tableColumn = columnModel.getColumn(i);
                    if (tableColumn.getMaxWidth() != 0) {
                        sheet.autoSizeColumn(i);
                    } else {
                        sheet.setColumnWidth(i, 0);
                    }
                }

                // Write to file
                try (FileOutputStream outputStream = new FileOutputStream(file)) {
                    workbook.write(outputStream);
                    JOptionPane.showMessageDialog(null,
                            "Watchlist exported successfully!",
                            "Export Complete",
                            JOptionPane.INFORMATION_MESSAGE);
                }
            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(null,
                        "Error exporting file: " + e.getMessage(),
                        "Error",
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }

// Helper methods to get column values for a specific row
    private double getPEAvgForRow(int row) {
        int column = findColumnByName("PE Avg");
        if (column != -1) {
            Object value = tableModel.getValueAt(row, column);
            if (value instanceof Number) {
                return ((Number) value).doubleValue();
            }
        }
        return 0.0;
    }

    private double getPBAvgForRow(int row) {
        int column = findColumnByName("PB Avg");
        if (column != -1) {
            Object value = tableModel.getValueAt(row, column);
            if (value instanceof Number) {
                return ((Number) value).doubleValue();
            }
        }
        return 0.0;
    }

    private double calculatePEForward(double price, double eps) {
        if (eps <= 0) {
            return 0.0;
        }
        return round(price / eps, 2);
    }

    private String[] getPEForwardColumnNames() {
        int currentYear = Calendar.getInstance().get(Calendar.YEAR);
        String[] peForwardColumnNames = new String[3];
        for (int i = 0; i < 3; i++) {
            peForwardColumnNames[i] = "PE FWD" + (currentYear + i);
        }
        return peForwardColumnNames;
    }

    private double getPriceToFCFAvgForRow(int row) {
        int column = findColumnByName("PFCF Avg");
        if (column != -1) {
            Object value = tableModel.getValueAt(row, column);
            if (value instanceof Number) {
                return ((Number) value).doubleValue();
            }
        }
        return 0.0;
    }

    private double getROEAvgForRow(int row) {
        int column = findColumnByName("ROE Avg");
        if (column != -1) {
            Object value = tableModel.getValueAt(row, column);
            if (value instanceof Number) {
                return ((Number) value).doubleValue();
            }
        }
        return 0.0;
    }

    private double getDEAvgForRow(int row) {
        int column = findColumnByName("DE Avg");
        if (column != -1) {
            Object value = tableModel.getValueAt(row, column);
            if (value instanceof Number) {
                return ((Number) value).doubleValue();
            }
        }
        return 0.0;
    }

    private int findColumnByName(String columnName) {
        for (int i = 0; i < tableModel.getColumnCount(); i++) {
            if (tableModel.getColumnName(i).equals(columnName)) {
                return i;
            }
        }
        return -1;
    }

    private double getPriceForRow(int row) {
        int priceColumn = -1;
        for (int i = 0; i < tableModel.getColumnCount(); i++) {
            if (tableModel.getColumnName(i).equals("Price")) {
                priceColumn = i;
                break;
            }
        }

        if (priceColumn != -1) {
            Object value = tableModel.getValueAt(row, priceColumn);
            if (value instanceof Number) {
                return ((Number) value).doubleValue();
            }
        }
        return 0.0;
    }

}
