package afin.jstocks;

import javax.swing.*;
import java.awt.*;
import java.util.ArrayList;

public class GUI {

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            // Create the main window
            JFrame mainFrame = new JFrame("Stock Manager");
            mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            mainFrame.setSize(800, 600);

            // Create tabs
            JTabbedPane tabbedPane = new JTabbedPane();

            // Add Watchlist Tab
            Watchlist watchlist = new Watchlist();
            tabbedPane.addTab("Watchlist", null, createTabPanel(() -> watchlist.createAndShowGUI()), "Manage your stock watchlist");

            // Add Portfolio Tab
            ArrayList<StockLot> stockLots = new ArrayList<>();
            Portfolio portfolio = new Portfolio(stockLots);
            tabbedPane.addTab("Portfolio", null, createTabPanel(() -> portfolio.createAndShowPortfolio()), "Manage your stock portfolio");

            // Add Analytics Tab (Placeholder for now)
            tabbedPane.addTab("Analytics", null, new JPanel(), "Analyze your stock data");

            // Add the tabbed pane to the main frame
            mainFrame.add(tabbedPane, BorderLayout.CENTER);

            // Display the main frame
            mainFrame.setLocationRelativeTo(null);
            mainFrame.setVisible(true);
        });
    }

    private static JPanel createTabPanel(Runnable tabInitializer) {
        JPanel panel = new JPanel(new BorderLayout());
        SwingUtilities.invokeLater(tabInitializer);
        return panel;
    }
}