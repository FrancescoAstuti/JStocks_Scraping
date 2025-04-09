package afin.jstocks;

import javax.swing.*;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableRowSorter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileWriter;
import java.io.IOException;
import java.io.FileReader;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Scanner;
import org.json.JSONArray;
import org.json.JSONObject;
import org.json.JSONTokener;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.time.Day;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;

public class StockLot {
    private String ticker;
    private double quantity;
    private double purchasePrice;
    private double currentPrice;
    private double volatilityScore; // Add volatility score field

    private static final String API_KEY = "eb7366217370656d66a56a057b8511b0";

    public StockLot(String ticker, double quantity, double purchasePrice, double currentPrice) {
        this.ticker = ticker;
        this.quantity = quantity;
        this.purchasePrice = purchasePrice;
        this.currentPrice = currentPrice;
        this.volatilityScore = calculateVolatilityScore(); // Calculate volatility score
    }

    public String getTicker() {
        return ticker;
    }

    public void setTicker(String ticker) {
        this.ticker = ticker;
    }

    public double getQuantity() {
        return quantity;
    }

    public void setQuantity(double quantity) {
        this.quantity = quantity;
    }

    public double getPurchasePrice() {
        return purchasePrice;
    }

    public void setPurchasePrice(double purchasePrice) {
        this.purchasePrice = purchasePrice;
    }

    public double getCurrentPrice() {
        return currentPrice;
    }

    public void setCurrentPrice(double currentPrice) {
        this.currentPrice = currentPrice;
    }

    public double getVolatilityScore() {
        return volatilityScore;
    }

    public double getProfitLossPercentage() {
        return ((currentPrice - purchasePrice) / purchasePrice) * 100;
    }

    public static double calculateTotalLots(List<StockLot> stockLots) {
        return stockLots.stream().mapToDouble(StockLot::getQuantity).sum();
    }

    public void updatePrice() throws IOException {
        String urlString = String.format("https://financialmodelingprep.com/api/v3/quote/%s?apikey=%s", ticker, API_KEY);
        HttpURLConnection connection = (HttpURLConnection) new URL(urlString).openConnection();
        connection.setRequestMethod("GET");

        int responseCode = connection.getResponseCode();
        if (responseCode == 200) {
            Scanner scanner = new Scanner(connection.getInputStream());
            String response = scanner.useDelimiter("\\A").next();
            scanner.close();

            JSONArray jsonArray = new JSONArray(response);
            if (jsonArray.length() > 0) {
                JSONObject jsonObject = jsonArray.getJSONObject(0);
                this.currentPrice = jsonObject.getDouble("price");
                this.volatilityScore = calculateVolatilityScore(); // Update volatility score
            }
        } else {
            throw new IOException("Failed to fetch stock price. Response code: " + responseCode);
        }
    }

    private double calculateVolatilityScore() {
        return Analytics.calculateVolatilityScore(ticker);
    }

    public static void createAndShowGUI(ArrayList<StockLot> stockLots) {
        JFrame frame = new JFrame("Stock Portfolio Manager");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);

        JPanel panel = new JPanel(new BorderLayout());

        DefaultTableModel tableModel = new DefaultTableModel(new Object[]{"Ticker", "Quantity", "Purchase Price", "Current Price", "P/L (%)", "Volatility Score"}, 0);
        JTable table = new JTable(tableModel);

        TableRowSorter<DefaultTableModel> sorter = new TableRowSorter<>(tableModel);
        sorter.setComparator(1, Comparator.comparingDouble(o -> Double.parseDouble(o.toString())));
        sorter.setComparator(2, Comparator.comparingDouble(o -> Double.parseDouble(o.toString())));
        sorter.setComparator(3, Comparator.comparingDouble(o -> Double.parseDouble(o.toString())));
        sorter.setComparator(4, Comparator.comparingDouble(o -> Double.parseDouble(o.toString())));
        sorter.setComparator(5, Comparator.comparingDouble(o -> Double.parseDouble(o.toString())));
        table.setRowSorter(sorter);

        JScrollPane scrollPane = new JScrollPane(table);
        panel.add(scrollPane, BorderLayout.CENTER);

        JPanel inputPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        JTextField tickerField = new JTextField();
        JTextField quantityField = new JTextField();
        JTextField purchasePriceField = new JTextField();

        Dimension preferredSize = new Dimension(200, 30);
        tickerField.setPreferredSize(preferredSize);
        quantityField.setPreferredSize(preferredSize);
        purchasePriceField.setPreferredSize(preferredSize);

        gbc.gridx = 0;
        gbc.gridy = 0;
        inputPanel.add(new JLabel("Ticker:"), gbc);
        gbc.gridx = 1;
        inputPanel.add(tickerField, gbc);

        gbc.gridx = 0;
        gbc.gridy = 1;
        inputPanel.add(new JLabel("Quantity:"), gbc);
        gbc.gridx = 1;
        inputPanel.add(quantityField, gbc);

        gbc.gridx = 0;
        gbc.gridy = 2;
        inputPanel.add(new JLabel("Purchase Price:"), gbc);
        gbc.gridx = 1;
        inputPanel.add(purchasePriceField, gbc);

        JButton addButton = new JButton("Add");
        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                int response = JOptionPane.showConfirmDialog(null, "Do you want to add this stock lot?", "Confirm", JOptionPane.YES_NO_OPTION);
                if (response == JOptionPane.YES_OPTION) {
                    addStockLot(stockLots, tableModel, tickerField, quantityField, purchasePriceField);
                }
            }
        });

        JButton modifyButton = new JButton("Modify");
        modifyButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                int response = JOptionPane.showConfirmDialog(null, "Do you want to modify this stock lot?", "Confirm", JOptionPane.YES_NO_OPTION);
                if (response == JOptionPane.YES_OPTION) {
                    modifyStockLot(stockLots, table, tableModel, tickerField, quantityField, purchasePriceField);
                }
            }
        });

        JButton deleteButton = new JButton("Delete");
        deleteButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                int response = JOptionPane.showConfirmDialog(null, "Do you want to delete this stock lot?", "Confirm", JOptionPane.YES_NO_OPTION);
                if (response == JOptionPane.YES_OPTION) {
                    deleteStockLot(stockLots, table, tableModel);
                }
            }
        });

        gbc.gridx = 0;
        gbc.gridy = 3;
        gbc.gridwidth = 2;
        inputPanel.add(addButton, gbc);
        gbc.gridy = 4;
        inputPanel.add(modifyButton, gbc);
        gbc.gridy = 5;
        inputPanel.add(deleteButton, gbc);

        panel.add(inputPanel, BorderLayout.SOUTH);

        frame.add(panel);
        frame.setVisible(true);

        table.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
            @Override
            public void valueChanged(ListSelectionEvent event) {
                if (!event.getValueIsAdjusting()) {
                    int viewRow = table.getSelectedRow();
                    if (viewRow != -1) {
                        int modelRow = table.convertRowIndexToModel(viewRow);
                        tickerField.setText((String) tableModel.getValueAt(modelRow, 0));
                        quantityField.setText(String.valueOf(tableModel.getValueAt(modelRow, 1)));
                        purchasePriceField.setText(String.valueOf(tableModel.getValueAt(modelRow, 2)));
                    }
                }
            }
        });

        updateTable(stockLots, tableModel, table);
    }

    private static void addStockLot(ArrayList<StockLot> stockLots, DefaultTableModel tableModel, JTextField tickerField, JTextField quantityField, JTextField purchasePriceField) {
        try {
            String ticker = tickerField.getText().toUpperCase();
            double quantity = Double.parseDouble(quantityField.getText());
            double purchasePrice = round(Double.parseDouble(purchasePriceField.getText()), 2);
            Double currentPrice = GetPrice.getCurrentPrice(ticker);

            if (currentPrice == null) {
                JOptionPane.showMessageDialog(null, "Ticker not found.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            currentPrice = round(currentPrice, 2);
            stockLots.add(new StockLot(ticker, quantity, purchasePrice, currentPrice));
            updateTable(stockLots, tableModel, null);
        } catch (NumberFormatException e) {
            JOptionPane.showMessageDialog(null, "Please enter valid numbers", "Error", JOptionPane.ERROR_MESSAGE);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private static void modifyStockLot(ArrayList<StockLot> stockLots, JTable table, DefaultTableModel tableModel, JTextField tickerField, JTextField quantityField, JTextField purchasePriceField) {
        int viewRow = table.getSelectedRow();
        if (viewRow != -1) {
            int modelRow = table.convertRowIndexToModel(viewRow);
            try {
                String ticker = tickerField.getText().toUpperCase();
                double quantity = Double.parseDouble(quantityField.getText());
                double purchasePrice = round(Double.parseDouble(purchasePriceField.getText()), 2);
                Double currentPrice = round(GetPrice.getCurrentPrice(ticker), 2);

                if (currentPrice == null) {
                    JOptionPane.showMessageDialog(null, "Failed to retrieve current price for ticker: " + ticker, "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                StockLot stockLot = stockLots.get(modelRow);
                stockLot.setTicker(ticker);
                stockLot.setQuantity(quantity);
                stockLot.setPurchasePrice(purchasePrice);
                stockLot.setCurrentPrice(currentPrice);
                stockLot.volatilityScore = stockLot.calculateVolatilityScore(); // Update volatility score
                updateTable(stockLots, tableModel, null);
            } catch (NumberFormatException e) {
                JOptionPane.showMessageDialog(null, "Please enter valid numbers", "Error", JOptionPane.ERROR_MESSAGE);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private static void deleteStockLot(ArrayList<StockLot> stockLots, JTable table, DefaultTableModel tableModel) {
        int viewRow = table.getSelectedRow();
        if (viewRow != -1) {
            int modelRow = table.convertRowIndexToModel(viewRow);
            stockLots.remove(modelRow);
            updateTable(stockLots, tableModel, null);
        }
    }

    private static void updateTable(ArrayList<StockLot> stockLots, DefaultTableModel tableModel, JTable table) {
        stockLots.sort(Comparator.comparing(StockLot::getTicker));

        tableModel.setRowCount(0);
        for (StockLot stockLot : stockLots) {
            tableModel.addRow(new Object[]{
                stockLot.getTicker().toUpperCase(),
                stockLot.getQuantity(),
                round(stockLot.getPurchasePrice(), 2),
                round(stockLot.getCurrentPrice(), 2),
                round(stockLot.getProfitLossPercentage(), 2),
                round(stockLot.getVolatilityScore(), 2) // Add volatility score to table
            });
        }

        if (table != null) {
            table.getColumnModel().getColumn(4).setCellRenderer(new DefaultTableCellRenderer() {
                @Override
                public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                    Component cell = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                    double plPercentage = (double) value;
                    if (plPercentage < -10) {
                        cell.setBackground(Color.RED);
                    } else if (plPercentage > 10) {
                        cell.setBackground(Color.GREEN);
                    } else {
                        cell.setBackground(Color.WHITE);
                    }
                    return cell;
                }
            });
        }
    }

    public static void saveStockLots(ArrayList<StockLot> stockLots) {
        JSONArray jsonArray = new JSONArray();
        for (StockLot stockLot : stockLots) {
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("ticker", stockLot.getTicker());
            jsonObject.put("quantity", stockLot.getQuantity());
            jsonObject.put("purchasePrice", stockLot.getPurchasePrice());
            jsonObject.put("currentPrice", stockLot.getCurrentPrice());
            jsonObject.put("volatilityScore", stockLot.getVolatilityScore()); // Save volatility score
            jsonArray.put(jsonObject);
        }

        try (FileWriter file = new FileWriter("stockLots.json")) {
            file.write(jsonArray.toString());
            System.out.println("Stock lots saved successfully.");
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Failed to save stock lots.");
        }
    }

    public static void loadStockLots(ArrayList<StockLot> stockLots) {
        try (FileReader reader = new FileReader("stockLots.json")) {
            JSONArray jsonArray = new JSONArray(new JSONTokener(reader));
            for (int i = 0; i < jsonArray.length(); i++) {
                JSONObject jsonObject = jsonArray.getJSONObject(i);
                String ticker = jsonObject.getString("ticker");
                double quantity = jsonObject.getDouble("quantity");
                double purchasePrice = jsonObject.getDouble("purchasePrice");
                double currentPrice = jsonObject.getDouble("currentPrice");
                double volatilityScore = jsonObject.getDouble("volatilityScore");
                StockLot stockLot = new StockLot(ticker, quantity, purchasePrice, currentPrice);
                stockLot.volatilityScore = volatilityScore; // Load volatility score
                stockLots.add(stockLot);
            }
            System.out.println("Stock lots loaded successfully.");
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Failed to load stock lots.");
        }
    }

    private static double round(double value, int places) {
        if (places < 0) throw new IllegalArgumentException();
        BigDecimal bd = BigDecimal.valueOf(value);
        bd = bd.setScale(places, RoundingMode.HALF_UP);
        return bd.doubleValue();
    }

    public static void main(String[] args) {
        ArrayList<StockLot> stockLots = new ArrayList<>();
        loadStockLots(stockLots);
        Runtime.getRuntime().addShutdownHook(new Thread(() -> saveStockLots(stockLots)));
        createAndShowGUI(stockLots);
    }
}