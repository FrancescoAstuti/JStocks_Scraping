package afin.jstocks;

import javax.swing.*;
import java.awt.*;
import java.util.ArrayList;

public class Main {
    private static ArrayList<StockLot> stockLots = new ArrayList<>();
    private static JFrame overviewFrame;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            showOverview();
        });
    }

    public static void showOverview() {
        if (overviewFrame != null) {
            overviewFrame.setVisible(true);
            return;
        }

        overviewFrame = new JFrame("Overview");
        overviewFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        overviewFrame.setSize(800, 600);

        // Add buttons for navigation
        JButton watchlistButton = new JButton("Watchlist");
        watchlistButton.addActionListener(e -> {
            overviewFrame.setVisible(false); // Hide overview window
            SwingUtilities.invokeLater(() -> {
                Watchlist watchlist = new Watchlist();
                watchlist.createAndShowGUI();
                overviewFrame.setVisible(true); // Restore overview window after Watchlist closes
            });
        });

        JButton portfolioButton = new JButton("Portfolio");
        portfolioButton.addActionListener(e -> {
            overviewFrame.setVisible(false); // Hide overview window
            SwingUtilities.invokeLater(() -> {
                Portfolio portfolio = new Portfolio(stockLots);
                portfolio.createAndShowPortfolio();
                overviewFrame.setVisible(true); // Restore overview window after Portfolio closes
            });
        });

        JButton stockScreenerButton = new JButton("Stock Screener");
        stockScreenerButton.addActionListener(e -> {
            StockScreener stockScreener = new StockScreener();
            stockScreener.createAndShowGUI();
        });

        // Main panel layout
        JPanel panel = new JPanel(new BorderLayout());
        panel.add(watchlistButton, BorderLayout.WEST);
        panel.add(portfolioButton, BorderLayout.CENTER);
        panel.add(stockScreenerButton, BorderLayout.EAST);

        overviewFrame.add(panel);
        overviewFrame.setLocationRelativeTo(null);
        overviewFrame.setVisible(true);
    }

    public void addStockToPortfolio(String ticker, double quantity, double purchasePrice, double currentPrice) {
        StockLot stockLot = new StockLot(ticker, quantity, purchasePrice, currentPrice);
        stockLots.add(stockLot);
    }
}