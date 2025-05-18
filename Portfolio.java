package afin.jstocks;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ArrayList;
import java.util.Comparator;
import org.json.JSONArray;
import org.json.JSONObject;
import org.json.JSONTokener;
import java.io.FileWriter;
import java.io.FileReader;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;

public class Portfolio {
    private ArrayList<StockLot> stockLots;
    private JTable table;
    private DefaultTableModel tableModel;
    private JTextField tickerField, quantityField, purchasePriceField;
    private JFrame portfolioFrame;

    public Portfolio(ArrayList<StockLot> stockLots) {
        this.stockLots = stockLots;
        loadStockLotsIfEmpty();
        Runtime.getRuntime().addShutdownHook(new Thread(this::saveStockLots));
    }

    public void createAndShowPortfolio() {
        portfolioFrame = new JFrame("Stock Portfolio Manager");
        portfolioFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE); // Dispose instead of exiting the program
        portfolioFrame.setSize(800, 600);

        JPanel panel = new JPanel(new BorderLayout());

        tableModel = new DefaultTableModel(new Object[]{"Ticker", "Quantity", "Purchase Price", "Current Price", "P/L (%)"}, 0);
        table = new JTable(tableModel);

        // Set custom cell renderer for the "P/L (%)" column
        table.getColumnModel().getColumn(4).setCellRenderer(new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                Component cell = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                if (value instanceof Double) {
                    double plPercentage = (Double) value;
                    if (plPercentage < -10) {
                        cell.setBackground(Color.RED);
                        cell.setForeground(Color.WHITE);
                    } else if (plPercentage > 10) {
                        cell.setBackground(Color.GREEN);
                        cell.setForeground(Color.BLACK);
                    } else {
                        cell.setBackground(Color.WHITE);
                        cell.setForeground(Color.BLACK);
                    }
                }
                return cell;
            }
        });

        JScrollPane scrollPane = new JScrollPane(table);
        panel.add(scrollPane, BorderLayout.CENTER);

        JPanel inputPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        tickerField = new JTextField();
        quantityField = new JTextField();
        purchasePriceField = new JTextField();

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
        addButton.addActionListener(e -> addStockLot());

        JButton modifyButton = new JButton("Modify");
        modifyButton.addActionListener(e -> modifyStockLot());

        JButton deleteButton = new JButton("Delete");
        deleteButton.addActionListener(e -> deleteStockLot());

        gbc.gridx = 0;
        gbc.gridy = 3;
        gbc.gridwidth = 2;
        inputPanel.add(addButton, gbc);
        gbc.gridy = 4;
        inputPanel.add(modifyButton, gbc);
        gbc.gridy = 5;
        inputPanel.add(deleteButton, gbc);

        panel.add(inputPanel, BorderLayout.SOUTH);

        portfolioFrame.add(panel);
        portfolioFrame.setLocationRelativeTo(null);
        portfolioFrame.setVisible(true);

        updateTable();
    }

    private void addStockLot() {
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
            updateTable();
        } catch (NumberFormatException e) {
            JOptionPane.showMessageDialog(null, "Please enter valid numbers", "Error", JOptionPane.ERROR_MESSAGE);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void modifyStockLot() {
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
                updateTable();
            } catch (NumberFormatException e) {
                JOptionPane.showMessageDialog(null, "Please enter valid numbers", "Error", JOptionPane.ERROR_MESSAGE);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void deleteStockLot() {
        int viewRow = table.getSelectedRow();
        if (viewRow != -1) {
            int modelRow = table.convertRowIndexToModel(viewRow);
            stockLots.remove(modelRow);
            updateTable();
        }
    }

    private void updateTable() {
        stockLots.sort(Comparator.comparing(StockLot::getTicker));

        tableModel.setRowCount(0);
        for (StockLot stockLot : stockLots) {
            tableModel.addRow(new Object[]{
                stockLot.getTicker().toUpperCase(),
                stockLot.getQuantity(),
                round(stockLot.getPurchasePrice(), 2),
                round(stockLot.getCurrentPrice(), 2),
                round(stockLot.getProfitLossPercentage(), 2)
            });
        }
    }

    private void saveStockLots() {
        JSONArray jsonArray = new JSONArray();
        for (StockLot stockLot : stockLots) {
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("ticker", stockLot.getTicker());
            jsonObject.put("quantity", stockLot.getQuantity());
            jsonObject.put("purchasePrice", stockLot.getPurchasePrice());
            jsonObject.put("currentPrice", stockLot.getCurrentPrice());
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

    private void loadStockLotsIfEmpty() {
        if (stockLots.isEmpty()) {
            loadStockLots();
        }
    }

    private void loadStockLots() {
        try (FileReader reader = new FileReader("stockLots.json")) {
            JSONArray jsonArray = new JSONArray(new JSONTokener(reader));
            for (int i = 0; i < jsonArray.length(); i++) {
                JSONObject jsonObject = jsonArray.getJSONObject(i);
                String ticker = jsonObject.getString("ticker");
                double quantity = jsonObject.getDouble("quantity");
                double purchasePrice = jsonObject.getDouble("purchasePrice");
                double currentPrice = jsonObject.getDouble("currentPrice");
                stockLots.add(new StockLot(ticker, quantity, purchasePrice, currentPrice));
            }
            System.out.println("Stock lots loaded successfully.");
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Failed to load stock lots.");
        }
    }

    private double round(double value, int places) {
        if (places < 0) throw new IllegalArgumentException();
        BigDecimal bd = BigDecimal.valueOf(value);
        bd = bd.setScale(places, RoundingMode.HALF_UP);
        return bd.doubleValue();
    }
}