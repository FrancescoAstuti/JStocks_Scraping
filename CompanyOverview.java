package afin.jstocks;

import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Scanner;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.ValueMarker;
import org.jfree.chart.ui.Layer;
import org.jfree.chart.ui.RectangleAnchor;
import org.jfree.chart.ui.TextAnchor;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;

import javax.swing.*;
import java.awt.*;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import org.json.JSONArray;
import org.json.JSONObject;

public class CompanyOverview {
    private static final Color BACKGROUND_COLOR = new Color(255, 255, 255);
    private static final Color GRID_LINE_COLOR = new Color(230, 230, 230);
    private static final Color BAR_COLOR = new Color(41, 128, 185);
    private static final Color TITLE_COLOR = new Color(44, 62, 80);
    private static final Color LABEL_COLOR = new Color(52, 73, 94);
    private static final Color MARKER_LINE_COLOR = new Color(231, 76, 60);
    private static final Color TTM_LINE_COLOR = new Color(46, 204, 113);
    private static final Font TITLE_FONT = new Font("Segoe UI", Font.BOLD, 16);
    private static final Font LABEL_FONT = new Font("Segoe UI", Font.PLAIN, 12);
    private static final String API_KEY = API_key.getApiKey();
    
public static String fetchIndustry(String ticker) {
    String urlString = String.format("https://financialmodelingprep.com/api/v3/profile/%s?apikey=%s", ticker, API_KEY);
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
                JSONObject profile = data.getJSONObject(0);
                return profile.optString("industry", "N/A");
            }
        }
        return "N/A";
    } catch (Exception e) {
        e.printStackTrace();
        return "N/A";
    } finally {
        if (connection != null) {
            connection.disconnect();
        }
    }
}    
    
private static Map<String, Double> getTTMValuesFromWatchlist(String ticker) {
    Map<String, Double> ttmValues = new HashMap<>();
    try {
        String jsonContent = new String(Files.readAllBytes(Paths.get("data/watchlist.json")));
        JSONObject watchlist = new JSONObject(jsonContent);
        JSONArray stocks = watchlist.getJSONArray("stocks");

        System.out.println("Looking for TTM values for ticker: " + ticker); // Debug print

        for (int i = 0; i < stocks.length(); i++) {
            JSONObject stock = stocks.getJSONObject(i);
            if (stock.getString("ticker").equals(ticker)) {
                ttmValues.put("peTTM", stock.getDouble("peTTM"));
                ttmValues.put("pbTTM", stock.getDouble("pbTTM"));
                ttmValues.put("epsTTM", stock.getDouble("epsTTM"));
                ttmValues.put("pfcfTTM", stock.getDouble("pfcfTTM"));
                ttmValues.put("debtToEquityTTM", stock.getDouble("debtToEquity"));
                ttmValues.put("roeTTM", stock.getDouble("roeTTM"));
                
                // Debug prints
                System.out.println("Found TTM values:");
                System.out.println("PE TTM: " + ttmValues.get("peTTM"));
                System.out.println("PB TTM: " + ttmValues.get("pbTTM"));
                System.out.println("EPS TTM: " + ttmValues.get("epsTTM"));
                System.out.println("PFCF TTM: " + ttmValues.get("pfcfTTM"));
                System.out.println("Debt to Equity TTM: " + ttmValues.get("debtToEquityTTM"));
                 System.out.println("ROE TTM: " + ttmValues.get("roeTTM"));
                break;
            }
        }
        
        if (ttmValues.isEmpty()) {
            System.out.println("No TTM values found for ticker: " + ticker); // Debug print
        }
    } catch (Exception e) {
        System.out.println("Error reading watchlist.json: " + e.getMessage()); // Debug print
        e.printStackTrace();
    }
    return ttmValues;
}

    public static void showCompanyOverview(String ticker, String companyName) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }

        JFrame overviewFrame = new JFrame("Company Overview: " + ticker);
        overviewFrame.setSize(1200, 1200);
        overviewFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);

        JPanel panel = new JPanel(new BorderLayout());
        JLabel titleLabel = new JLabel("Company Name: " + companyName, SwingConstants.CENTER);
        titleLabel.setFont(TITLE_FONT);
        titleLabel.setForeground(TITLE_COLOR);
        panel.add(titleLabel, BorderLayout.NORTH);

        // Fetch historical data
        List<RatioData> peRatios = Ratios.fetchHistoricalPE(ticker);
        List<RatioData> pbRatios = Ratios.fetchHistoricalPB(ticker);
        List<RatioData> epsRatios = Ratios.fetchQuarterlyEPS(ticker);
        List<RatioData> pfcfRatios = Ratios.fetchHistoricalPFCF(ticker);
        List<RatioData> debtToEquityRatios = Ratios.fetchHistoricalDebtToEquity(ticker);
        List<RatioData> roeRatios = Ratios.fetchHistoricalReturnOnEquity(ticker);

        // Filter and sort data
        LocalDate timeRange = LocalDate.now().minusYears(20);
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

        peRatios = filterAndSortRatios(peRatios, timeRange, formatter);
        pbRatios = filterAndSortRatios(pbRatios, timeRange, formatter);
        epsRatios = filterAndSortRatios(epsRatios, timeRange, formatter);
        pfcfRatios = filterAndSortRatios(pfcfRatios, timeRange, formatter);
        debtToEquityRatios = filterAndSortRatios(debtToEquityRatios, timeRange, formatter);
        roeRatios = filterAndSortRatios(roeRatios, timeRange, formatter);

        // Calculate averages
        List<Double> cappedPeValues = capValues(peRatios, 0.0, 40.0);
        List<Double> cappedPfcfValues = capValues(pfcfRatios, -50.0, 50.0);
        List<Double> cappedRoeValues = capValues(roeRatios, -50.0, 50.0);

        double peAverage = calculateAverage(cappedPeValues);
        double pbAverage = calculateAverage(pbRatios.stream()
                .map(RatioData::getValue).collect(Collectors.toList()));
        double epsAverage = calculateAverage(epsRatios.stream()
                .map(RatioData::getValue).collect(Collectors.toList()));
        double pfcfAverage = calculateAverage(cappedPfcfValues);
        double debtToEquityAverage = calculateAverage(debtToEquityRatios.stream()
                .map(RatioData::getValue).collect(Collectors.toList()));
        double roeAverage = calculateAverage(cappedRoeValues);

        // Get TTM values from watchlist
        Map<String, Double> ttmValues = getTTMValuesFromWatchlist(ticker);
        // Create datasets
        DefaultCategoryDataset peDataset = createDataset(peRatios, "PE Ratio", -40.0, 40.0);
        DefaultCategoryDataset pbDataset = createDataset(pbRatios, "PB Ratio", -10.0, 10.0);
        DefaultCategoryDataset epsDataset = createDataset(epsRatios, "EPS", null, null);
        DefaultCategoryDataset pfcfDataset = createDataset(pfcfRatios, "PFCF Ratio", -50.0, 50.0);
        DefaultCategoryDataset debtToEquityDataset = createDataset(debtToEquityRatios, "Debt to Equity Ratio", null, null);
        DefaultCategoryDataset roeDataset = createDataset(roeRatios, "ROE", -50.0, 50.0);

        // Create charts
        JFreeChart peChart = createStyledChart("PE Ratio", "Date", "PE Ratio", peDataset);
        JFreeChart pbChart = createStyledChart("PB Ratio", "Date", "PB Ratio", pbDataset);
        JFreeChart epsChart = createStyledChart("EPS", "Date", "EPS", epsDataset);
        JFreeChart pfcfChart = createStyledChart("PFCF Ratio", "Date", "PFCF Ratio", pfcfDataset);
        JFreeChart debtToEquityChart = createStyledChart("Debt to Equity", 
                "Date", "Debt to Equity Ratio", debtToEquityDataset);
        JFreeChart roeChart = createStyledChart("Return on Equity", "Date", "ROE", roeDataset);

        // Add markers for averages and TTM values
        addMarker(peChart, peAverage, "Average PE", MARKER_LINE_COLOR);
        addMarker(peChart, ttmValues.getOrDefault("peTTM", Double.NaN), "TTM PE", TTM_LINE_COLOR);

        addMarker(pbChart, pbAverage, "Average PB", MARKER_LINE_COLOR);
        addMarker(pbChart, ttmValues.getOrDefault("pbTTM", Double.NaN), "TTM PB", TTM_LINE_COLOR);

        addMarker(epsChart, epsAverage, "Average EPS", MARKER_LINE_COLOR);
        addMarker(epsChart, ttmValues.getOrDefault("epsTTM", Double.NaN), "TTM EPS", TTM_LINE_COLOR);

        addMarker(pfcfChart, pfcfAverage, "Average PFCF", MARKER_LINE_COLOR);
        addMarker(pfcfChart, ttmValues.getOrDefault("pfcfTTM", Double.NaN), "TTM PFCF", TTM_LINE_COLOR);

        addMarker(debtToEquityChart, debtToEquityAverage, "Average Debt to Equity", MARKER_LINE_COLOR);
        addMarker(debtToEquityChart, ttmValues.getOrDefault("debtToEquityTTM", Double.NaN), 
                "TTM Debt to Equity", TTM_LINE_COLOR);
        
        addMarker(roeChart, roeAverage, "Average ROE", MARKER_LINE_COLOR);
        addMarker(roeChart, ttmValues.getOrDefault("roeTTM", Double.NaN), "TTM ROE", TTM_LINE_COLOR);

        // Setup panel for charts
        JPanel chartsPanel = new JPanel(new GridLayout(6, 1, 0, 10));
        chartsPanel.setBackground(BACKGROUND_COLOR);
        chartsPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // Add charts to panel
        addChartToPanel(chartsPanel, peChart);
        addChartToPanel(chartsPanel, pbChart);
        addChartToPanel(chartsPanel, epsChart);
        addChartToPanel(chartsPanel, pfcfChart);
        addChartToPanel(chartsPanel, debtToEquityChart);
        addChartToPanel(chartsPanel, roeChart);

        // Create scroll pane
        JScrollPane scrollPane = new JScrollPane(chartsPanel);
        scrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        scrollPane.getViewport().setBackground(BACKGROUND_COLOR);

        panel.add(scrollPane, BorderLayout.CENTER);
        panel.setBackground(BACKGROUND_COLOR);

        overviewFrame.add(panel);
        overviewFrame.setLocationRelativeTo(null);
        overviewFrame.setVisible(true);

        // Save data
        saveDataToFile(ticker, peRatios, pbRatios, epsRatios, pfcfRatios, debtToEquityRatios, roeRatios,
                peAverage, pbAverage, epsAverage, pfcfAverage, debtToEquityAverage, roeAverage);
    }

    private static List<RatioData> filterAndSortRatios(List<RatioData> ratios, LocalDate timeRange, 
            DateTimeFormatter formatter) {
        return ratios.stream()
                .filter(data -> LocalDate.parse(data.getDate(), formatter).isAfter(timeRange))
                .sorted(Comparator.comparing(data -> LocalDate.parse(data.getDate(), formatter)))
                .collect(Collectors.toList());
    }

    private static List<Double> capValues(List<RatioData> ratios, Double min, Double max) {
        return ratios.stream()
                .map(RatioData::getValue)
                .map(value -> {
                    if (min != null && max != null) {
                        return Math.max(Math.min(value, max), min);
                    }
                    return value;
                })
                .collect(Collectors.toList());
    }

    private static double calculateAverage(List<Double> values) {
        return values.stream()
                .filter(value -> value > 0)
                .mapToDouble(Double::doubleValue)
                .average()
                .orElse(Double.NaN);
    }

    private static DefaultCategoryDataset createDataset(List<RatioData> ratios, String series, 
            Double min, Double max) {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        for (RatioData data : ratios) {
            double value = data.getValue();
            if (min != null && max != null) {
                value = Math.max(Math.min(value, max), min);
            }
            dataset.addValue(value, series, data.getDate());
        }
        return dataset;
    }
    private static JFreeChart createStyledChart(String title, String xLabel, String yLabel, 
            DefaultCategoryDataset dataset) {
        JFreeChart chart = ChartFactory.createBarChart(
                title, xLabel, yLabel,
                dataset, PlotOrientation.VERTICAL,
                true, true, false);
        styleChart(chart);
        return chart;
    }

    private static void styleChart(JFreeChart chart) {
        org.jfree.chart.plot.CategoryPlot plot = chart.getCategoryPlot();
        
        // Background and grid
        chart.setBackgroundPaint(BACKGROUND_COLOR);
        plot.setBackgroundPaint(BACKGROUND_COLOR);
        plot.setRangeGridlinePaint(GRID_LINE_COLOR);
        plot.setDomainGridlinePaint(GRID_LINE_COLOR);
        
        // Bar styling
        BarRenderer renderer = (BarRenderer) plot.getRenderer();
        renderer.setBarPainter(new StandardBarPainter());
        renderer.setShadowVisible(false);
        renderer.setMaximumBarWidth(0.1);
        
        // Gradient paint
        GradientPaint gradientPaint = new GradientPaint(
            0.0f, 0.0f, BAR_COLOR,
            0.0f, 0.0f, BAR_COLOR.darker()
        );
        renderer.setSeriesPaint(0, gradientPaint);

        // Axis styling
        CategoryAxis domainAxis = plot.getDomainAxis();
        domainAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_45);
        domainAxis.setTickLabelFont(LABEL_FONT);
        domainAxis.setLabelFont(LABEL_FONT);
        domainAxis.setTickLabelPaint(LABEL_COLOR);
        domainAxis.setLabelPaint(LABEL_COLOR);

        org.jfree.chart.axis.ValueAxis rangeAxis = plot.getRangeAxis();
        rangeAxis.setTickLabelFont(LABEL_FONT);
        rangeAxis.setLabelFont(LABEL_FONT);
        rangeAxis.setTickLabelPaint(LABEL_COLOR);
        rangeAxis.setLabelPaint(LABEL_COLOR);

        // Title and legend styling
        chart.getTitle().setFont(TITLE_FONT);
        chart.getTitle().setPaint(TITLE_COLOR);
        chart.getLegend().setItemFont(LABEL_FONT);
        chart.getLegend().setBackgroundPaint(BACKGROUND_COLOR);
    }

   private static void addMarker(JFreeChart chart, double value, String label, Color color) {
    System.out.println("Adding marker: " + label + " with value: " + value); // Debug print
    if (!Double.isNaN(value)) {
        ValueMarker marker = new ValueMarker(value);
        marker.setLabel(String.format("%s: %.2f", label, value));
        marker.setLabelAnchor(RectangleAnchor.TOP_RIGHT);
        marker.setLabelTextAnchor(TextAnchor.BOTTOM_RIGHT);
        marker.setLabelFont(LABEL_FONT);
        marker.setPaint(color);
        marker.setStroke(new BasicStroke(1.0f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER, 
                1.0f, new float[] {2.0f, 2.0f}, 0.0f));
        chart.getCategoryPlot().addRangeMarker(marker, Layer.FOREGROUND);
        System.out.println("Marker added successfully"); // Debug print
    } else {
        System.out.println("Skipping marker due to NaN value"); // Debug print
    }
}

    private static void addChartToPanel(JPanel panel, JFreeChart chart) {
        ChartPanel chartPanel = new ChartPanel(chart);
        chartPanel.setPreferredSize(new Dimension(1100, 400));
        chartPanel.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createEmptyBorder(5, 5, 5, 5),
            BorderFactory.createLineBorder(GRID_LINE_COLOR)
        ));
        panel.add(chartPanel);
    }

    private static void saveDataToFile(String ticker, List<RatioData> peRatios, List<RatioData> pbRatios,
            List<RatioData> epsRatios, List<RatioData> pfcfRatios, List<RatioData> debtToEquityRatios, List<RatioData> roeRatios,
            double peAverage, double pbAverage, double epsAverage, double pfcfAverage, 
            double debtToEquityAverage, double roeAverage) {
        JSONObject jsonObject = new JSONObject();
        
        jsonObject.put("ticker", ticker);
        jsonObject.put("lastUpdated", LocalDate.now().toString());
        jsonObject.put("peAverage", peAverage);
        jsonObject.put("pbAverage", pbAverage);
        jsonObject.put("epsAverage", epsAverage);
        jsonObject.put("pfcfAverage", pfcfAverage);
        jsonObject.put("debtToEquityAverage", debtToEquityAverage);
        jsonObject.put("roeAverage", roeAverage);

        addRatiosToJson(jsonObject, "peRatios", peRatios, 0.0, 40.0);
        addRatiosToJson(jsonObject, "pbRatios", pbRatios, null, null);
        addRatiosToJson(jsonObject, "epsRatios", epsRatios, null, null);
        addRatiosToJson(jsonObject, "pfcfRatios", pfcfRatios, -50.0, 50.0);
        addRatiosToJson(jsonObject, "debtToEquityRatios", debtToEquityRatios, null, null);
        addRatiosToJson(jsonObject, "roeRatios", roeRatios, -50.0, 50.0);

        try (FileWriter file = new FileWriter("data/" + ticker + "_overview.json")) {
            file.write(jsonObject.toString(2));
            file.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void addRatiosToJson(JSONObject jsonObject, String key, List<RatioData> ratios, 
            Double min, Double max) {
        JSONArray ratioArray = new JSONArray();
        for (RatioData data : ratios) {
            JSONObject dataObject = new JSONObject();
            dataObject.put("date", data.getDate());
            
            double value = data.getValue();
            if (min != null && max != null) {
                value = Math.max(Math.min(value, max), min);
            }
            dataObject.put("value", value);
            
            ratioArray.put(dataObject);
        }
        jsonObject.put(key, ratioArray);
    }
}