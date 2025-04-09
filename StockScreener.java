package afin.jstocks;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import org.json.JSONArray;
import org.json.JSONObject;

public class StockScreener {
    private JTextField marketCapMinField;
    private JTextField marketCapMaxField;
    private JTextField dividendYieldMinField;
    private JTextField dividendYieldMaxField;
    private JTextField pegRatioMinField;
    private JTextField pegRatioMaxField;
    private JTable resultTable;
    private DefaultTableModel tableModel;

    private static final String API_KEY = "eb7366217370656d66a56a057b8511b0";

    public void createAndShowGUI() {
        JFrame frame = new JFrame("Stock Screener");
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        frame.setSize(1000, 800);

        JPanel mainPanel = new JPanel(new BorderLayout());

        JPanel settingsPanel = new JPanel(new GridLayout(8, 2));
        settingsPanel.setPreferredSize(new Dimension(300, 800));

        settingsPanel.add(new JLabel("Market Cap Min:"));
        marketCapMinField = new JTextField();
        marketCapMinField.setPreferredSize(new Dimension(100, 20));
        settingsPanel.add(marketCapMinField);

        settingsPanel.add(new JLabel("Market Cap Max:"));
        marketCapMaxField = new JTextField();
        marketCapMaxField.setPreferredSize(new Dimension(100, 20));
        settingsPanel.add(marketCapMaxField);

        settingsPanel.add(new JLabel("Dividend Yield Min:"));
        dividendYieldMinField = new JTextField();
        dividendYieldMinField.setPreferredSize(new Dimension(100, 20));
        settingsPanel.add(dividendYieldMinField);

        settingsPanel.add(new JLabel("Dividend Yield Max:"));
        dividendYieldMaxField = new JTextField();
        dividendYieldMaxField.setPreferredSize(new Dimension(100, 20));
        settingsPanel.add(dividendYieldMaxField);

        settingsPanel.add(new JLabel("PEG Ratio Min:"));
        pegRatioMinField = new JTextField();
        pegRatioMinField.setPreferredSize(new Dimension(100, 20));
        settingsPanel.add(pegRatioMinField);

        settingsPanel.add(new JLabel("PEG Ratio Max:"));
        pegRatioMaxField = new JTextField();
        pegRatioMaxField.setPreferredSize(new Dimension(100, 20));
        settingsPanel.add(pegRatioMaxField);

        JButton searchButton = new JButton("Search");
        searchButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                searchStocks();
            }
        });
        settingsPanel.add(searchButton);

        JButton overviewButton = new JButton("Return to Overview");
        overviewButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                frame.dispose();
                Main.showOverview();
            }
        });
        settingsPanel.add(overviewButton);

        mainPanel.add(settingsPanel, BorderLayout.WEST);

        tableModel = new DefaultTableModel(new Object[]{"Name", "Ticker", "Capitalization", "Dividend Yield", "PEG Ratio"}, 0);
        resultTable = new JTable(tableModel);
        JScrollPane scrollPane = new JScrollPane(resultTable);
        mainPanel.add(scrollPane, BorderLayout.CENTER);

        frame.add(mainPanel);
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }

    private void searchStocks() {
        String marketCapMin = marketCapMinField.getText();
        String marketCapMax = marketCapMaxField.getText();
        String dividendYieldMin = dividendYieldMinField.getText();
        String dividendYieldMax = dividendYieldMaxField.getText();
        String pegRatioMin = pegRatioMinField.getText();
        String pegRatioMax = pegRatioMaxField.getText();

        List<String[]> filteredStocks = fetchFilteredStocks(marketCapMin, marketCapMax, dividendYieldMin, dividendYieldMax, pegRatioMin, pegRatioMax);

        tableModel.setRowCount(0);
        for (String[] stock : filteredStocks) {
            tableModel.addRow(stock);
        }
    }

    private List<String[]> fetchFilteredStocks(String marketCapMin, String marketCapMax, String dividendYieldMin, String dividendYieldMax, String pegRatioMin, String pegRatioMax) {
        List<String[]> stocks = new ArrayList<>();
        try {
            String urlString = String.format("https://financialmodelingprep.com/api/v3/stock-screener?marketCapMoreThan=%s&marketCapLessThan=%s&dividendMoreThan=%s&dividendLessThan=%s&pegMoreThan=%s&pegLessThan=%s&apikey=%s",
                    marketCapMin, marketCapMax, dividendYieldMin, dividendYieldMax, pegRatioMin, pegRatioMax, API_KEY);

            URL url = new URL(urlString);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");

            BufferedReader in = new BufferedReader(new InputStreamReader(connection.getInputStream()));
            String inputLine;
            StringBuilder content = new StringBuilder();
            while ((inputLine = in.readLine()) != null) {
                content.append(inputLine);
            }
            in.close();

            JSONArray jsonArray = new JSONArray(content.toString());
            for (int i = 0; i < jsonArray.length(); i++) {
                JSONObject jsonObject = jsonArray.getJSONObject(i);
                String name = jsonObject.getString("companyName");
                String ticker = jsonObject.getString("symbol");
                String capitalization = String.valueOf(jsonObject.getLong("marketCap"));

                // Fetch key metrics for each ticker
                String keyMetricsUrl = String.format("https://financialmodelingprep.com/api/v3/key-metrics-ttm/%s?apikey=%s", ticker, API_KEY);
                URL keyMetricsEndpoint = new URL(keyMetricsUrl);
                HttpURLConnection keyMetricsConnection = (HttpURLConnection) keyMetricsEndpoint.openConnection();
                keyMetricsConnection.setRequestMethod("GET");

                BufferedReader keyMetricsIn = new BufferedReader(new InputStreamReader(keyMetricsConnection.getInputStream()));
                StringBuilder keyMetricsContent = new StringBuilder();
                while ((inputLine = keyMetricsIn.readLine()) != null) {
                    keyMetricsContent.append(inputLine);
                }
                keyMetricsIn.close();

                JSONArray keyMetricsArray = new JSONArray(keyMetricsContent.toString());
                JSONObject keyMetricsObject = keyMetricsArray.getJSONObject(0);

                // Print the fetched JSON object for debugging
                System.out.println("JSON response for " + ticker + ": " + keyMetricsObject.toString());

                String dividendYield = String.valueOf(keyMetricsObject.optDouble("dividendYield", 0.0));
                String pegRatio = String.valueOf(keyMetricsObject.optDouble("pegRatio", 0.0));

                // Print fetched data to the console
                System.out.println("Fetched data for " + ticker + ":");
                System.out.println("Name: " + name);
                System.out.println("Capitalization: " + capitalization);
                System.out.println("Dividend Yield: " + dividendYield);
                System.out.println("PEG Ratio: " + pegRatio);

                stocks.add(new String[]{name, ticker, capitalization, dividendYield, pegRatio});
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return stocks;
    }
}