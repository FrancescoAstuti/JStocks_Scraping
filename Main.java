package afin.jstocks;

import org.jfree.chart.ChartPanel;
import javax.swing.*;
import java.awt.BorderLayout;
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

        GUI gui = new GUI(stockLots);
        ChartPanel chartPanel = gui.createChartPanel();

        JButton portfolioButton = new JButton("Portfolio");
        portfolioButton.addActionListener(e -> {
            overviewFrame.dispose();
            gui.createAndShowGUI();
        });

        JButton stockScreenerButton = new JButton("Stock Screener");
        stockScreenerButton.addActionListener(e -> {
            StockScreener stockScreener = new StockScreener();
            stockScreener.createAndShowGUI();
        });

        JButton watchlistButton = new JButton("Watchlist");
        watchlistButton.addActionListener(e -> {
            Watchlist watchlist = new Watchlist();
            watchlist.createAndShowGUI();
        });

        JPanel timeFramePanel = new JPanel();
        String[] timeFrames = {"1W", "1M", "2M", "3M", "6M", "1Y"};
        for (String timeFrame : timeFrames) {
            JButton button = new JButton(timeFrame);
            button.addActionListener(e -> gui.updateChart(timeFrame));
            timeFramePanel.add(button);
        }

        JPanel panel = new JPanel(new BorderLayout());
        panel.add(timeFramePanel, BorderLayout.NORTH);
        panel.add(chartPanel, BorderLayout.CENTER);
        panel.add(portfolioButton, BorderLayout.SOUTH);
        panel.add(stockScreenerButton, BorderLayout.EAST);
        panel.add(watchlistButton, BorderLayout.WEST);

        overviewFrame.add(panel);
        overviewFrame.setLocationRelativeTo(null);
        overviewFrame.setVisible(true);
    }

    public void addStockToPortfolio(String ticker, double quantity, double purchasePrice, double currentPrice) {
        StockLot stockLot = new StockLot(ticker, quantity, purchasePrice, currentPrice);
        stockLots.add(stockLot);
    }
}