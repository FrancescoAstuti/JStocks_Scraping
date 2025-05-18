package afin.jstocks;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import java.awt.*;
import java.awt.Color;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import java.io.FileOutputStream;

public class CellsFormat {

    // Soft base colors
    private static final Color LIGHT_RED = new Color(255, 235, 235);
    private static final Color MEDIUM_RED = new Color(255, 200, 200); // New
    private static final Color DARK_RED = new Color(255, 180, 180);   // New
    private static final Color LIGHT_YELLOW = new Color(255, 255, 220);

    // Soft Graham Number colors
    private static final Color LIGHT_GREEN = new Color(220, 255, 220);    // Very soft green
    private static final Color MEDIUM_GREEN = new Color(198, 255, 198);   // Soft mint green
    private static final Color DARK_GREEN = new Color(178, 255, 178);     // Pastel green

    private static final Color LIGHT_PINK = new Color(255, 230, 230);     // Very soft pink
    private static final Color MEDIUM_PINK = new Color(255, 200, 200);    // Soft pink
    private static final Color DARK_PINK = new Color(255, 180, 180);      // Pastel pink

    public static class CustomCellRenderer extends DefaultTableCellRenderer {

        @Override
        public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
            Component cell = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

            if (!isSelected) { // Don't change background if cell is selected
                String columnName = table.getColumnName(column);

                if (columnName.equals("Graham Number") && value instanceof Double) {
                    double grahamNumber = (Double) value;
                    // Get the price from the "Price" column
                    int priceColumn = -1;
                    for (int i = 0; i < table.getColumnCount(); i++) {
                        if (table.getColumnName(i).equals("Price")) {
                            priceColumn = i;
                            break;
                        }
                    }

                    if (priceColumn != -1) {
                        Object priceObj = table.getValueAt(row, priceColumn);
                        if (priceObj instanceof Double) {
                            double price = (Double) priceObj;
                            if (price > 0) {
                                double percentDiff = (grahamNumber - price) / price * 100;

                                if (percentDiff > 0) {
                                    if (percentDiff <= 25) {
                                        cell.setBackground(LIGHT_GREEN);
                                    } else if (percentDiff <= 50) {
                                        cell.setBackground(MEDIUM_GREEN);
                                    } else {
                                        cell.setBackground(DARK_GREEN);
                                    }
                                } else {
                                    percentDiff = Math.abs(percentDiff);
                                    if (percentDiff <= 25) {
                                        cell.setBackground(LIGHT_PINK);
                                    } else if (percentDiff <= 50) {
                                        cell.setBackground(MEDIUM_PINK);
                                    } else {
                                        cell.setBackground(DARK_PINK);
                                    }
                                }
                                cell.setForeground(new Color(51, 51, 51));
                                return cell;
                            }
                        }
                    }
                } else if (columnName.equals("Debt to Equity") && value instanceof Double) {
                    double debtToEquity = (Double) value;
                    int deAvgColumn = -1;
                    for (int i = 0; i < table.getColumnCount(); i++) {
                        if (table.getColumnName(i).equals("DE Avg")) {
                            deAvgColumn = i;
                            break;
                        }
                    }

                    if (deAvgColumn != -1) {
                        Object deAvgObj = table.getValueAt(row, deAvgColumn);
                        if (deAvgObj instanceof Double) {
                            double deAvg = (Double) deAvgObj;
                            if (deAvg > 0 && debtToEquity > 0) {
                                double ratio = debtToEquity / deAvg;
                                if (ratio < 1) {
                                    if (ratio >= 0.75) cell.setBackground(LIGHT_GREEN);
                                    else if (ratio >= 0.5) cell.setBackground(MEDIUM_GREEN);
                                    else cell.setBackground(DARK_GREEN);
                                } else {
                                    if (ratio <= 1.25) cell.setBackground(LIGHT_PINK);
                                    else if (ratio <= 1.5) cell.setBackground(MEDIUM_PINK);
                                    else cell.setBackground(DARK_PINK);
                                }
                                cell.setForeground(new Color(51, 51, 51));
                                return cell;
                            }
                        }
                    }
                } else if (columnName.equals("PE TTM") && value instanceof Double) {
                    double peTtm = (Double) value;
                    int peAvgColumn = -1;
                    for (int i = 0; i < table.getColumnCount(); i++) {
                        if (table.getColumnName(i).equals("PE Avg")) {
                            peAvgColumn = i;
                            break;
                        }
                    }
                    if (peAvgColumn != -1) {
                        Object peAvgObj = table.getValueAt(row, peAvgColumn);
                        if (peAvgObj instanceof Double) {
                            double peAvg = (Double) peAvgObj;
                            if (peAvg > 0 && peTtm > 0) {
                                double ratio = peTtm / peAvg;
                                if (ratio < 1) {
                                    if (ratio >= 0.75) cell.setBackground(LIGHT_GREEN);
                                    else if (ratio >= 0.5) cell.setBackground(MEDIUM_GREEN);
                                    else cell.setBackground(DARK_GREEN);
                                } else {
                                    if (ratio <= 1.25) cell.setBackground(LIGHT_PINK);
                                    else if (ratio <= 1.5) cell.setBackground(MEDIUM_PINK);
                                    else cell.setBackground(DARK_PINK);
                                }
                                cell.setForeground(new Color(51, 51, 51));
                                return cell;
                            }
                        }
                    }
                } else if (columnName.startsWith("PE FWD") && value instanceof Double) {
                    double peForward = (Double) value;
                    int peAvgColumn = -1;
                    for (int i = 0; i < table.getColumnCount(); i++) {
                        if (table.getColumnName(i).equals("PE Avg")) {
                            peAvgColumn = i;
                            break;
                        }
                    }
                    if (peAvgColumn != -1) {
                        Object peAvgObj = table.getValueAt(row, peAvgColumn);
                        if (peAvgObj instanceof Double) {
                            double peAvg = (Double) peAvgObj;
                            if (peAvg > 0 && peForward > 0) {
                                double ratio = peForward / peAvg;
                                if (ratio < 1) {
                                    if (ratio >= 0.75) cell.setBackground(LIGHT_GREEN);
                                    else if (ratio >= 0.5) cell.setBackground(MEDIUM_GREEN);
                                    else cell.setBackground(DARK_GREEN);
                                } else {
                                    if (ratio <= 1.25) cell.setBackground(LIGHT_PINK);
                                    else if (ratio <= 1.5) cell.setBackground(MEDIUM_PINK);
                                    else cell.setBackground(DARK_PINK);
                                }
                                cell.setForeground(new Color(51, 51, 51));
                                return cell;
                            }
                        }
                    }
                } else if (columnName.equals("PB TTM") && value instanceof Double) {
                    double pbTtm = (Double) value;
                    int pbAvgColumn = -1;
                    for (int i = 0; i < table.getColumnCount(); i++) {
                        if (table.getColumnName(i).equals("PB Avg")) {
                            pbAvgColumn = i;
                            break;
                        }
                    }
                    if (pbAvgColumn != -1) {
                        Object pbAvgObj = table.getValueAt(row, pbAvgColumn);
                        if (pbAvgObj instanceof Double) {
                            double pbAvg = (Double) pbAvgObj;
                            if (pbAvg > 0 && pbTtm > 0) {
                                double ratio = pbTtm / pbAvg;
                                if (ratio < 1) {
                                    if (ratio >= 0.75) cell.setBackground(LIGHT_GREEN);
                                    else if (ratio >= 0.5) cell.setBackground(MEDIUM_GREEN);
                                    else cell.setBackground(DARK_GREEN);
                                } else {
                                    if (ratio <= 1.25) cell.setBackground(LIGHT_PINK);
                                    else if (ratio <= 1.5) cell.setBackground(MEDIUM_PINK);
                                    else cell.setBackground(DARK_PINK);
                                }
                                cell.setForeground(new Color(51, 51, 51));
                                return cell;
                            }
                        }
                    }
                } else if (columnName.equals("P/FCF") && value instanceof Double) {
                    double pfcfTtm = (Double) value;
                    int pfcfAvgColumn = -1;
                    for (int i = 0; i < table.getColumnCount(); i++) {
                        if (table.getColumnName(i).equals("PFCF Avg")) {
                            pfcfAvgColumn = i;
                            break;
                        }
                    }
                    if (pfcfAvgColumn != -1) {
                        Object pfcfAvgObj = table.getValueAt(row, pfcfAvgColumn);
                        if (pfcfAvgObj instanceof Double) {
                            double pfcfAvg = (Double) pfcfAvgObj;
                            if (pfcfAvg > 0 && pfcfTtm > 0) {
                                double ratio = pfcfTtm / pfcfAvg;
                                if (ratio < 1) {
                                    if (ratio >= 0.75) cell.setBackground(LIGHT_GREEN);
                                    else if (ratio >= 0.5) cell.setBackground(MEDIUM_GREEN);
                                    else cell.setBackground(DARK_GREEN);
                                } else {
                                    if (ratio <= 1.25) cell.setBackground(LIGHT_PINK);
                                    else if (ratio <= 1.5) cell.setBackground(MEDIUM_PINK);
                                    else cell.setBackground(DARK_PINK);
                                }
                                cell.setForeground(new Color(51, 51, 51));
                                return cell;
                            }
                        }
                    }
                } else if (columnName.equals("ROE TTM") && value instanceof Double) {
                    double roeTtm = (Double) value;
                    int roeAvgColumn = -1;
                    for (int i = 0; i < table.getColumnCount(); i++) {
                        if (table.getColumnName(i).equals("ROE Avg")) {
                            roeAvgColumn = i;
                            break;
                        }
                    }
                    if (roeAvgColumn != -1) {
                        Object roeAvgObj = table.getValueAt(row, roeAvgColumn);
                        if (roeAvgObj instanceof Double) {
                            double roeAvg = (Double) roeAvgObj;
                            if (roeAvg > 0 && roeTtm > 0) {
                                double ratio = roeTtm / roeAvg;
                                if (ratio > 1) {
                                    if (ratio <= 1.25) cell.setBackground(LIGHT_GREEN);
                                    else if (ratio <= 1.5) cell.setBackground(MEDIUM_GREEN);
                                    else cell.setBackground(DARK_GREEN);
                                } else {
                                    if (ratio >= 0.75) cell.setBackground(LIGHT_PINK);
                                    else if (ratio >= 0.5) cell.setBackground(MEDIUM_PINK);
                                    else cell.setBackground(DARK_PINK);
                                }
                                cell.setForeground(new Color(51, 51, 51));
                                return cell;
                            }
                        }
                    }
                } else if (columnName.equals("Payout Ratio") && value instanceof Double) {
                    double payoutRatio = (Double) value;

                    // New Payout Ratio formatting logic
                    if (payoutRatio < 0) {
                        cell.setBackground(DARK_RED);       // Dark Red: negative values
                    } else if (payoutRatio == 0) {
                        cell.setBackground(LIGHT_YELLOW);   // Light Yellow: payoutRatio == 0
                    } else if (payoutRatio > 0 && payoutRatio <= 0.33) {
                        cell.setBackground(DARK_GREEN);     // Dark Green: 0% < payoutRatio <= 33%
                    } else if (payoutRatio > 0.33 && payoutRatio <= 0.66) {
                        cell.setBackground(MEDIUM_GREEN);   // Medium Green: 33% < payoutRatio <= 66%
                    } else if (payoutRatio > 0.66 && payoutRatio < 1.00) { // 66% < payoutRatio < 100%
                        cell.setBackground(LIGHT_GREEN);    // Light Green
                    } else if (payoutRatio >= 1.00 && payoutRatio <= 1.25) { // 100% <= payoutRatio <= 125%
                        cell.setBackground(LIGHT_RED);      // Light Red
                    } else if (payoutRatio > 1.25 && payoutRatio <= 1.50) { // 125% < payoutRatio <= 150%
                        cell.setBackground(MEDIUM_RED);     // Medium Red
                    } else if (payoutRatio > 1.50) {
                        cell.setBackground(DARK_RED);       // Dark Red: payoutRatio > 150%
                    } else {
                        // Default for any other cases (e.g., NaN, though unlikely if instanceof Double)
                        cell.setBackground(Color.WHITE);
                    }
                    cell.setForeground(new Color(51, 51, 51));
                    return cell;

                } else if ((columnName.equals("EPS Growth 1")
                        || columnName.equals("EPS Growth 2")
                        || columnName.equals("EPS Growth 3"))
                        && value instanceof Double) {

                    double epsGrowth = (Double) value;

                    if (epsGrowth < 0) {
                        cell.setBackground(LIGHT_RED);
                    } else if (epsGrowth < 15) {
                        cell.setBackground(LIGHT_GREEN);
                    } else if (epsGrowth < 30) {
                        cell.setBackground(MEDIUM_GREEN);
                    } else {
                        cell.setBackground(DARK_GREEN);
                    }
                    cell.setForeground(new Color(51, 51, 51));
                    return cell;
                }

                // Default coloring for other cells
                if (value instanceof Double) {
                    double numValue = (Double) value;
                    if (numValue < 0) {
                        cell.setBackground(LIGHT_RED);
                    } else if (numValue == 0.0) {
                        cell.setBackground(LIGHT_YELLOW);
                    } else {
                        cell.setBackground(Color.WHITE);
                    }
                } else if ("n/a".equals(value)) {
                    cell.setBackground(Color.LIGHT_GRAY);
                } else {
                    cell.setBackground(Color.WHITE);
                }
                cell.setForeground(Color.BLACK);
            }
            return cell;
        }
    }

    public static XSSFCellStyle createCustomColorStyle(XSSFWorkbook workbook, byte[] rgb) {
        XSSFCellStyle style = workbook.createCellStyle();
        XSSFColor color = new XSSFColor(rgb, null); // For XSSF (xlsx)
        // For HSSF (xls), you might need to find the closest matching HSSFColor or use a palette.
        // However, since you are using XSSFWorkbook, XSSFColor is correct.
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
}