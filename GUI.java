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
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
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

public class GUI {
    private ArrayList<StockLot> stockLots;
    private JTable table;
    private DefaultTableModel tableModel;
    private JTextField tickerField, quantityField, purchasePriceField;
    private TimeSeries series;
    private ChartPanel chartPanel;
    private JFreeChart timeChart;

    public GUI(ArrayList<StockLot> stockLots) {
        this.stockLots = stockLots;
        loadStockLotsIfEmpty();
        Runtime.getRuntime().addShutdownHook(new Thread(this::saveStockLots));
    }

    public void createAndShowGUI() {
        JFrame frame = new JFrame("Stock Portfolio Manager");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);

        JPanel panel = new JPanel(new BorderLayout());

        tableModel = new DefaultTableModel(new Object[]{"Ticker", "Quantity", "Purchase Price", "Current Price", "P/L (%)"}, 0);
        table = new JTable(tableModel);

        // Enable sorting with custom comparators
        TableRowSorter<DefaultTableModel> sorter = new TableRowSorter<>(tableModel);
        sorter.setComparator(1, Comparator.comparingDouble(o -> Double.parseDouble(o.toString()))); // Quantity column
        sorter.setComparator(2, Comparator.comparingDouble(o -> Double.parseDouble(o.toString()))); // Purchase Price column
        sorter.setComparator(3, Comparator.comparingDouble(o -> Double.parseDouble(o.toString()))); // Current Price column
        sorter.setComparator(4, Comparator.comparingDouble(o -> Double.parseDouble(o.toString()))); // P/L (%) column
        table.setRowSorter(sorter);

        JScrollPane scrollPane = new JScrollPane(table);
        panel.add(scrollPane, BorderLayout.CENTER);

        JPanel inputPanel = new JPanel(new GridBagLayout()); // Use GridBagLayout for better control
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        tickerField = new JTextField();
        quantityField = new JTextField();
        purchasePriceField = new JTextField();

        // Set preferred size for the text fields
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
                    addStockLot();
                }
            }
        });

        JButton modifyButton = new JButton("Modify");
        modifyButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                int response = JOptionPane.showConfirmDialog(null, "Do you want to modify this stock lot?", "Confirm", JOptionPane.YES_NO_OPTION);
                if (response == JOptionPane.YES_OPTION) {
                    modifyStockLot();
                }
            }
        });

        JButton deleteButton = new JButton("Delete");
        deleteButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                int response = JOptionPane.showConfirmDialog(null, "Do you want to delete this stock lot?", "Confirm", JOptionPane.YES_NO_OPTION);
                if (response == JOptionPane.YES_OPTION) {
                    deleteStockLot();
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

        // Add a button to return to the overview window
        JButton overviewButton = new JButton("Overview");
        overviewButton.addActionListener(e -> {
            frame.dispose();
            Main.showOverview();
        });
        gbc.gridy = 6;
        inputPanel.add(overviewButton, gbc);

        panel.add(inputPanel, BorderLayout.SOUTH);

        frame.add(panel);
        frame.setVisible(true);

        // Add a ListSelectionListener to the table to handle row selection
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

        // Call updateTable to ensure the stock list is displayed immediately
        updateTable();
    }

    public ChartPanel createChartPanel() {
        series = new TimeSeries("Total Value");
        timeChart = ChartFactory.createTimeSeriesChart(
            "Total Value Over Time",
            "Date",
            "Total Value",
            new TimeSeriesCollection(series),
            true,
            true,
            false
        );
        chartPanel = new ChartPanel(timeChart);
        return chartPanel;
    }

    public void updateChart(String timeFrame) {
        series.clear();
        LocalDate endDate = LocalDate.now();
        LocalDate startDate;
        switch (timeFrame) {
            case "1W":
                startDate = endDate.minus(1, ChronoUnit.WEEKS);
                break;
            case "1M":
                startDate = endDate.minus(1, ChronoUnit.MONTHS);
                break;
            case "2M":
                startDate = endDate.minus(2, ChronoUnit.MONTHS);
                break;
            case "3M":
                startDate = endDate.minus(3, ChronoUnit.MONTHS);
                break;
            case "6M":
                startDate = endDate.minus(6, ChronoUnit.MONTHS);
                break;
            case "1Y":
            default:
                startDate = endDate.minus(1, ChronoUnit.YEARS);
                break;
        }

        for (LocalDate date = startDate; !date.isAfter(endDate); date = date.plusDays(1)) {
            double totalValue = stockLots.stream()
                .mapToDouble(stockLot -> stockLot.getQuantity() * stockLot.getCurrentPrice())
                .sum();
            series.add(new Day(date.getDayOfMonth(), date.getMonthValue(), date.getYear()), totalValue);
        }

        timeChart.getXYPlot().setDataset(new TimeSeriesCollection(series));
        chartPanel.repaint();
    }

    private void addStockLot() {
        try {
            String ticker = tickerField.getText().toUpperCase();
            double quantity = Double.parseDouble(quantityField.getText()); // Changed from int to double
            double purchasePrice = round(Double.parseDouble(purchasePriceField.getText()), 2);
            Double currentPrice = GetPrice.getCurrentPrice(ticker);

            if (currentPrice == null) {
                JOptionPane.showMessageDialog(null, "Ticker not found.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            currentPrice = round(currentPrice, 2);
            stockLots.add(new StockLot(ticker, quantity, purchasePrice, currentPrice)); // Changed quantity type to double
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
                double quantity = Double.parseDouble(quantityField.getText()); // Changed from int to double
                double purchasePrice = round(Double.parseDouble(purchasePriceField.getText()), 2);
                Double currentPrice = round(GetPrice.getCurrentPrice(ticker), 2);

                if (currentPrice == null) {
                    JOptionPane.showMessageDialog(null, "Failed to retrieve current price for ticker: " + ticker, "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                StockLot stockLot = stockLots.get(modelRow);
                stockLot.setTicker(ticker);
                stockLot.setQuantity(quantity); // Changed quantity type to double
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
        // Sort the underlying data
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

        // Set custom cell renderer for P/L (%) column
        table.getColumnModel().getColumn(4).setCellRenderer(new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                Component cell = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                double plPercentage = (double) value;
                if (plPercentage < -10) {
                    cell.setBackground(Color.RED);
                } 
                else if (plPercentage > 10) {
                    cell.setBackground(Color.GREEN);
                }
                else {
                    cell.setBackground(Color.WHITE);
                }
                
                            
                return cell;
            }
        });
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
                double quantity = jsonObject.getDouble("quantity"); // Changed from int to double
                double purchasePrice = jsonObject.getDouble("purchasePrice");
                double currentPrice = jsonObject.getDouble("currentPrice");
                stockLots.add(new StockLot(ticker, quantity, purchasePrice, currentPrice)); // Changed quantity type to double
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