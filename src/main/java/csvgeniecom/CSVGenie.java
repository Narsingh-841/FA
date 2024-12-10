package csvgeniecom;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;

import java.io.*;
import java.nio.file.*;
import java.text.DecimalFormat;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

class XLSXRecord {
	 private String assetNumber;                      // Updated from "code"
	    private String assetName;                        // Updated from "name"
	    private String purchaseDate;                     // New field
	    private String purchasePrice;                    // New field
	    private String assetType;                        // Updated from "type"
	    private String bookRate;                         // New field
	    private String closingBookRate;                  // New field
	    private String bookAccumulatedDepreciation;      // New field
	    private String bookDepMethodValue;              // New field
	    private String taxRate;
	    private String taxDepMethodValue;   
	    public String getTaxAccumulatedDepreciation() {
			return TaxAccumulatedDepreciation;
		}

		public void setTaxAccumulatedDepreciation(String taxAccumulatedDepreciation) {
			TaxAccumulatedDepreciation = taxAccumulatedDepreciation;
		}

		private String TaxAccumulatedDepreciation;
	    public String getTaxDepMethodValue() {
			return taxDepMethodValue;
		}

		public void setTaxDepMethodValue(String taxDepMethodValue) {
			this.taxDepMethodValue = taxDepMethodValue;
		}

		public String getTaxRate() {
			return taxRate;
		}

		public void setTaxRate(String taxRate) {
			this.taxRate = taxRate;
		}

		// Getters and Setters
	    public String getAssetNumber() {
	        return assetNumber;
	    }

	    public void setAssetNumber(String assetNumber) {
	        this.assetNumber = assetNumber;
	    }

	    public String getAssetName() {
	        return assetName;
	    }

	    public void setAssetName(String assetName) {
	        this.assetName = assetName;
	    }

	    public String getPurchaseDate() {
	        return purchaseDate;
	    }

	    public void setPurchaseDate(String purchaseDate) {
	        this.purchaseDate = purchaseDate;
	    }

	    public String getPurchasePrice() {
	        return purchasePrice;
	    }

	    public void setPurchasePrice(String purchasePrice) {
	        this.purchasePrice = purchasePrice;
	    }

	    public String getAssetType() {
	        return assetType;
	    }

	    public void setAssetType(String assetType) {
	        this.assetType = assetType;
	    }

	    public String getBookRate() {
	        return bookRate;
	    }

	    public void setBookRate(String bookRate) {
	        this.bookRate = bookRate;
	    }

	    public String getClosingBookRate() {
	        return closingBookRate;
	    }

	    public void setClosingBookRate(String closingBookRate) {
	        this.closingBookRate = closingBookRate;
	    }

	    public String getBookAccumulatedDepreciation() {
	        return bookAccumulatedDepreciation;
	    }

	    public void setBookAccumulatedDepreciation(String bookAccumulatedDepreciation) {
	        this.bookAccumulatedDepreciation = bookAccumulatedDepreciation;
	    }

	    public String getBookDepMethodValue() {
	        return bookDepMethodValue;
	    }

	    public void setBookDepMethodValue(String bookDepMethodValue) {
	        this.bookDepMethodValue = bookDepMethodValue;
	    }
}


public class CSVGenie {

    public static void main(String[] args) {
        SwingUtilities.invokeLater(CSVGenie::createAndShowGUI);
    }

    private static void createAndShowGUI() {
        JFrame frame = new JFrame("FA2");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(500, 400);
        frame.setResizable(false);

        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BorderLayout());
        mainPanel.setBackground(Color.WHITE);
        mainPanel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));

        JLabel headerLabel = new JLabel("FA2", JLabel.CENTER);
        headerLabel.setFont(new Font("Arial", Font.BOLD, 24));
        headerLabel.setForeground(new Color(50, 120, 200));

        JPanel centerPanel = new JPanel();
        centerPanel.setLayout(new GridBagLayout());
        centerPanel.setBackground(Color.WHITE);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(10, 10, 10, 10);

        JLabel instructions = new JLabel("<html>NOTE:Please Ensure that column names are placed correctly <br> 1. Select multiple XLSX files.<br>2. Choose the output directory.</html>");
        instructions.setFont(new Font("Arial", Font.PLAIN, 14));
        instructions.setForeground(Color.DARK_GRAY);

        JButton chooseFilesButton = new JButton("Choose Excel Files");
        JButton chooseOutputButton = new JButton("Choose Output Directory");
        JButton convertButton = new JButton("Convert to CSV");
        convertButton.setBackground(new Color(50, 120, 200));
        convertButton.setForeground(Color.WHITE);

        JLabel statusLabel = new JLabel("Status: Waiting for user input.", JLabel.CENTER);
        statusLabel.setFont(new Font("Arial", Font.ITALIC, 12));
        statusLabel.setForeground(new Color(100, 100, 100));

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setMultiSelectionEnabled(true);
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));

        JFileChooser outputChooser = new JFileChooser();
        outputChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        List<File> selectedFiles = new ArrayList<>();
        File[] outputDirectory = new File[1];

        chooseFilesButton.addActionListener(e -> {
            int returnValue = fileChooser.showOpenDialog(frame);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                selectedFiles.clear();
                selectedFiles.addAll(Arrays.asList(fileChooser.getSelectedFiles()));
                statusLabel.setText("Selected " + selectedFiles.size() + " files.");
            }
        });

        chooseOutputButton.addActionListener(e -> {
            int returnValue = outputChooser.showOpenDialog(frame);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                outputDirectory[0] = outputChooser.getSelectedFile();
                statusLabel.setText("Output directory set to: " + outputDirectory[0].getAbsolutePath());
            }
        });

        convertButton.addActionListener(e -> {
            if (selectedFiles.isEmpty() || outputDirectory[0] == null) {
                JOptionPane.showMessageDialog(frame, "Please select files and an output directory.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            JDialog progressDialog = new JDialog(frame, "Processing...", true);
            progressDialog.setSize(300, 100);
            progressDialog.setLayout(new BorderLayout());
            progressDialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
         
            JLabel progressLabel = new JLabel("Converting files, please wait...", JLabel.CENTER);
            progressDialog.add(progressLabel, BorderLayout.CENTER);
            progressDialog.setLocationRelativeTo(frame);

            SwingWorker<Void, Void> worker = new SwingWorker<>() {
                @Override
                protected Void doInBackground() throws InterruptedException {
                    try {
                    	 Thread.sleep(5000);
                        for (File file : selectedFiles) {
                            String outputFilePath = new File(outputDirectory[0], file.getName().replace(".xlsx", ".csv")).getAbsolutePath();
                            List<XLSXRecord> records = readXLSXFile(file.getAbsolutePath());
                            writeToCSV(records, outputFilePath);
                        }
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(frame, "An error occurred: The column name or the following column data is not present" );
                    }
                    return null;
                }

                @Override
                protected void done() {
                    progressDialog.dispose();
                    statusLabel.setText("Conversion complete! Files saved in: " + outputDirectory[0].getAbsolutePath());
                }
            };

            worker.execute();
            progressDialog.setVisible(true);
        });

        gbc.gridx = 0;
        gbc.gridy = 0;
        centerPanel.add(instructions, gbc);
        gbc.gridy++;
        centerPanel.add(chooseFilesButton, gbc);
        gbc.gridy++;
        centerPanel.add(chooseOutputButton, gbc);
        gbc.gridy++;
        centerPanel.add(convertButton, gbc);

        mainPanel.add(headerLabel, BorderLayout.NORTH);
        mainPanel.add(centerPanel, BorderLayout.CENTER);
        mainPanel.add(statusLabel, BorderLayout.SOUTH);

        frame.add(mainPanel);
        frame.setVisible(true);
    }

    public static List<XLSXRecord> readXLSXFile(String filePath) throws IOException {
        List<XLSXRecord> records = new ArrayList<>();
        String lastAssetType = null;

        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) { // Use XSSFWorkbook for .xlsx files
            Sheet sheet = workbook.getSheetAt(0); // Read the first sheet

            // Read the header row and map column names to indices
            Row headerRow = sheet.getRow(7); // Assuming the header is in the 10th row (index 9)
            Map<String, Integer> columnIndices = getColumnIndices(headerRow);

            for (int i = 6; i <= sheet.getLastRowNum(); i++) { // Start iterating from row 11 (index 10)
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Debug: Print row index and first cell value
                String firstCellValue = getCellValue(row.getCell(0));
//                System.out.println("Processing row " + i + " (First Cell: " + firstCellValue + ")");

                // Check if the row contains "Grand Total"
                if (firstCellValue != null && firstCellValue.trim().equalsIgnoreCase("Grand Total")) {
                    break; // Stop processing further rows
                }

                // Skip rows if any cell in columns A to M (0 to 12) is empty
                if (isRowCompletelyEmptyInAMColumns(row)) {
                    continue; // Skip the row
                }

                XLSXRecord record = new XLSXRecord();

                // Read data dynamically using column indices
                Integer assetCodeIndex = columnIndices.get("Asset         Code");
                if (assetCodeIndex != null) {
                    record.setAssetNumber(getCellValue(row.getCell(assetCodeIndex)));
                } else {
                    // Handle missing column case (e.g., log an error or set a default value)
                    System.err.println("Column 'Asset         Code' not found.");
                }
 // AssetNumber
                record.setAssetName(getCellValue(row.getCell(columnIndices.get("Description"))));    // AssetName
                record.setPurchaseDate(getCellValue(row.getCell(columnIndices.get("Acq. (Disp.)         Date")))); // PurchaseDate
                Integer purchasepriceIndex = columnIndices.get("Original         Cost");
                if (purchasepriceIndex != null) {
                    record.setPurchasePrice(getCellValue(row.getCell(purchasepriceIndex)));
                } else {
                    // Handle missing column case (e.g., log an error or set a default value)
                    System.err.println("Column 'Asset         Code' not found.");
                }// PurchasePrice
                record.setBookRate(getCellValue(row.getCell(columnIndices.get("Dep.          Rate %"))));    // BookRate
                record.setBookAccumulatedDepreciation(getCellValue(row.getCell(columnIndices.get("Closing          W.D.V"))));
                record.setTaxRate(getCellValue(row.getCell(columnIndices.get("Effective        Life or %"))));
record.setTaxAccumulatedDepreciation(getCellValue(row.getCell(columnIndices.get("Closing Adj. Value"))));
               
                // Set the depreciation method based on the BookRate
                String bookRate = record.getBookRate();
                String bookDepMethodValue = "";

                if (bookRate != null) {
                    if (bookRate.contains("DV")) {
                        bookDepMethodValue = "Diminishing Value";
                    } else if (bookRate.contains("SL") || bookRate.contains("PC")) {
                        bookDepMethodValue = "Straight Line";
                    } else if (bookRate.contains("IWO")) {
                        bookDepMethodValue = "Full depreciation at purchase";
                    }else 
                        bookDepMethodValue = "Diminishing Value";
                }
                record.setBookDepMethodValue(bookDepMethodValue); // Set the BookDepreciationMethod
                // Set the depreciation method based on the TaxRate
                String taxRate = record.getBookRate();
                String taxDepMethodValue = "";

                if (taxRate != null) {
                    if (taxRate.contains("DV")) {
                    	taxDepMethodValue = "Diminishing Value";
                    } else if (taxRate.contains("SL") || bookRate.contains("PC")) {
                    	taxDepMethodValue = "Straight Line";
                    } else if (taxRate.contains("IWO")) {
                    	taxDepMethodValue = "Full depreciation at purchase";
                    }else
                    	taxDepMethodValue = "Diminishing Value";
                }
                record.setTaxDepMethodValue(taxDepMethodValue);
                String assetNumber = record.getAssetNumber();
                String assetName = record.getAssetName();

                // Update AssetType only for meaningful rows
                if (isCellBold(row.getCell(columnIndices.get("Description")))) {
                    // Case 1: AssetNumber length is 3
                    String assetType = assetName != null && !assetName.isEmpty() ? assetName : ""; // Use assetName or empty string
                    record.setAssetType(assetType); // Set the AssetType
                    lastAssetType = assetType; // Update the last valid AssetType
                } else if (!isCellBold(row.getCell(7))) {
                    // Case 2: AssetNumber is valid but not 3 digits
                    record.setAssetType(lastAssetType); // Use the last valid AssetType
                } else {
                    // Case 3: Empty or invalid AssetNumber
                    record.setAssetType(""); // Fallback to empty string
                }

                // Skip rows where all key fields are empty
                if ((assetNumber == null || assetNumber.isEmpty()) &&
                    (assetName == null || assetName.isEmpty()) &&
                    (record.getPurchaseDate() == null || record.getPurchaseDate().isEmpty()) &&
                    (record.getPurchasePrice() == null || record.getPurchasePrice().isEmpty())) {
                    continue; // Skip this row
                }

                // Debug: Print the values being read for the current row
//                System.out.println("AssetNumber: " + record.getAssetNumber() + ", AssetName: " + record.getAssetName() +
//                        ", PurchaseDate: " + record.getPurchaseDate() + ", PurchasePrice: " + record.getPurchasePrice() +
//                        ", AssetType: " + record.getAssetType() + ", BookRate: " + record.getBookRate() +
//                        ", BookDepreciationMethod: " + record.getBookDepMethodValue() + ", BookAccumulatedDepreciation: " + record.getBookAccumulatedDepreciation()+ ", taxrate: " + record.getTaxRate()+ ", taxDepmethod: " + record.getTaxDepMethodValue()+ ", TaxAccumulatedDepreciation: " + record.getTaxAccumulatedDepreciation());

                records.add(record);
            }
        }catch (IOException ex) {
            JFrame errorFrame = new JFrame(); // Create a new frame for the error dialog
            JOptionPane.showMessageDialog(
            	    errorFrame, 
            	    "An error occurred: Check if the column name is placed in the correct column data. If not, please solve and retry.", 
            	    "Error", 
            	    JOptionPane.ERROR_MESSAGE
            	);
//            System.out.println(ex.getMessage());
        }catch (Exception ex) {
            JFrame errorFrame = new JFrame(); // Create a new frame for the error dialog
            JOptionPane.showMessageDialog(
            	    errorFrame, 
            	    "An error occurred: Check if the column name is placed in the correct column data. If not, please solve and retry.", 
            	    "Error", 
            	    JOptionPane.ERROR_MESSAGE
            	);
//            System.out.println(ex.getMessage());
        }
        return records;
    }
    private static Map<String, Integer> getColumnIndices(Row headerRow) {
        Map<String, Integer> columnIndices = new HashMap<>();
        for (Cell cell : headerRow) {
            String columnName = getCellValue(cell).trim();
            columnIndices.put(columnName, cell.getColumnIndex());
        }
        // Debug: Print all column names and their indices
//        System.out.println("Header Columns: " + columnIndices);
        return columnIndices;
    }


    private static boolean isCellBold(Cell cell) {
        if (cell == null) return false;
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle == null) return false;
        org.apache.poi.ss.usermodel.Font font = cell.getSheet().getWorkbook().getFontAt(cellStyle.getFontIndexAsInt());
        return font.getBold(); // Returns true if the font is bold
    }


 // Helper method to check if any row is completely empty from columns A to M
    private static boolean isRowCompletelyEmptyInAMColumns(Row row) {
        // Iterate through columns A to M (indices 0 to 12)
        for (int j = 0; j <= 12; j++) {  // Columns A to M (indices 0 to 12)
            Cell cell = row.getCell(j); // Get the cell at column index j
            if (cell != null && !getCellValue(cell).trim().isEmpty()) {
                // If any cell in columns A to M has data, return false (do not skip this row)
                return false; 
            }
        }
        // If all cells in columns A to M are empty, return true (skip this row)
        return true; 
    }

 // Helper method to check if a string is numeric
    private static boolean isNumeric(String str) {
        if (str == null || str.isEmpty() ) {
            return false;
        }
        try {
            Double.parseDouble(str);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
    

    public static void writeToCSV(List<XLSXRecord> records, String outputFilePath) throws IOException {
        try (BufferedWriter bw = Files.newBufferedWriter(Paths.get(outputFilePath))) {
            // Write the header row
        	 String header = "*AssetName, *AssetNumber, PurchaseDate, PurchasePrice, AssetType,"
                     + "Description, TrackingCategory1, TrackingOption1, TrackingCategory2, TrackingOption2,"
                     + "SerialNumber, WarrantyExpiry, Book_DepreciationStartDate, Book_CostLimit, Book_ResidualValue,"
                     + "Book_DepreciationMethod, Book_AveragingMethod, Book_Rate, Book_EffectiveLife, Book_OpeningBookAccumulatedDepreciation,"
                     + "Tax_DepreciationMethod, Tax_PoolName, Tax_PooledDate, Tax_PooledAmount, Tax_DepreciationStartDate,"
                     + "Tax_CostLimit, Tax_ResidualValue, Tax_AveragingMethod, Tax_Rate, Tax_EffectiveLife, Tax_OpeningAccumulatedDepreciation";
            bw.write(header);
            bw.newLine();
            // Call formatDate() to format the purchaseDate
            // Write each record
            for (XLSXRecord record : records) {
            	String purchaseDate = record.getPurchaseDate();

// Skip rows where PurchaseDate is empty
            	if (purchaseDate != null && (purchaseDate.contains("Acq. (Disp.)         Date") || purchaseDate.trim().isEmpty())) {
            	    continue;
            	}
                if (purchaseDate != null && purchaseDate.contains("(") && purchaseDate.contains(")")) {
                    purchaseDate += "notice";
                }
                record.setPurchaseDate(purchaseDate);
                // Extract only numeric characters from BookRate
                String bookRate = record.getBookRate();
                String numericBookRate = bookRate != null ? bookRate.replaceAll("[^\\d.]", "") : "";
                
             // Extract only numeric characters from taxRate
                String taxRate = record.getTaxRate();
                String numericTaxRate = taxRate != null ? taxRate.replaceAll("[^\\d.]", "") : "";
                
//                double numericBookRateValue = Double.parseDouble(numericBookRate); // Convert String to double
                String bookAccumulatedDepreciation = record.getBookAccumulatedDepreciation();
                double accumulatedDepreciationValue = 0.0;
                if (bookAccumulatedDepreciation != null && !bookAccumulatedDepreciation.trim().isEmpty()) {
                    bookAccumulatedDepreciation = bookAccumulatedDepreciation.replaceAll(",", ""); // Remove commas
                    if (isNumeric(bookAccumulatedDepreciation)) {
                        accumulatedDepreciationValue = Double.parseDouble(bookAccumulatedDepreciation);
                    }else {
                    	accumulatedDepreciationValue=0.0;
                    }
                }
                // Calculate Tax Opening Accumulated Depreciation
                String purchasePriceStr = record.getPurchasePrice().replaceAll(",", "");
                double purchasePrice = 0.0;
                if (purchasePriceStr != null && isNumeric(purchasePriceStr)) {
                    purchasePrice = Double.parseDouble(purchasePriceStr.replaceAll("[^\\d.]", ""));
                }
                double BookOpeningAccumulatedDepreciation = purchasePrice - accumulatedDepreciationValue;
                
                String taxAccumulatedDepreciation = record.getTaxAccumulatedDepreciation();
                double taxAccumulatedDepreciationValue = 0.0;
                if (taxAccumulatedDepreciation != null && !taxAccumulatedDepreciation.trim().isEmpty()) {
                    taxAccumulatedDepreciation = taxAccumulatedDepreciation.replaceAll(",", ""); // Remove commas
                    if (isNumeric(taxAccumulatedDepreciation)) {
                        taxAccumulatedDepreciationValue = Double.parseDouble(taxAccumulatedDepreciation);
                    }else {
                    	taxAccumulatedDepreciationValue = 0.0; 
                    }
                double taxOpeningAccumulatedDepreciation = purchasePrice - taxAccumulatedDepreciationValue;

                
//                System.out.println("purchasePrice"+purchasePrice+"taxAccumulatedDepreciationValue"+taxAccumulatedDepreciationValue);
                

                
                
                

                String row = String.join(",",
                        "\"" + cleanValue(record.getAssetName()) + "\"", //AssetName
                        "\"" + cleanValue(record.getAssetNumber()) + "\"",//AssetNumber
                        "\"" + cleanValue(record.getPurchaseDate()) + "\"",//PurchaseDate
                		 "\"" + cleanValue(record.getPurchasePrice()) + "\"",//PurchasePrice
                "\"" + cleanValue(record.getAssetType()) + "\"",//AssetType
                "\" \"",//Description
                "\" \"",//TrackingCategory1
                "\" \"",//TrackingOption1
                "\" \"",//TrackingCategory2
                "\" \"",//TrackingOption2
                "\" \"",//SerialNumber
                "\" \"",//WarrantyExpiry
                "\"" + cleanValue(record.getPurchaseDate()) + "\"",//Book_DepreciationStartDate
                "\" \"",//Book_CostLimit
                "\" \"",//Book_ResidualValue
                "\"" + cleanValue(record.getBookDepMethodValue()) + "\"",//Book_DepreciationMethod
                "\"Actual Days \"",//Book_AveragingMethod
                "\"" + numericBookRate  + "\"", //Book_Rate
                "\" \"",//Book_EffectiveLife
                "\"" + BookOpeningAccumulatedDepreciation + "\"",//Book_OpeningBookAccumulatedDepreciation
                "\"" + cleanValue(record.getTaxDepMethodValue()) + "\"",//Tax_DepreciationMethod
                "\" \"",//Tax_PoolName
                "\" \"",//Tax_PooledDate
                "\" \"",//Tax_PooledAmount
                "\"" + cleanValue(record.getPurchaseDate()) + "\"",//Tax_DepreciationStartDate
                "\" \"",//Tax_CostLimit
                "\" \"",//Tax_ResidualValue
                "\"Actual Days \"",//Tax_AveragingMethod
                "\"" + numericTaxRate + "\"",//Tax_Rate
                "\" \"",//Tax_EffectiveLife
                "\"" + taxOpeningAccumulatedDepreciation + "\"");
//Tax_OpeningAccumulatedDepreciation
                
                	
                
                

                bw.write(row);
                bw.newLine();
            }}}catch (IOException ex) {
            JFrame errorFrame = new JFrame(); // Create a new frame for the error dialog
            JOptionPane.showMessageDialog(
            	    errorFrame, 
            	    "An error occurred: Check if the column name is placed in the correct column data. If not, please solve and retry.", 
            	    "Error", 
            	    JOptionPane.ERROR_MESSAGE
            	);
            System.out.println(ex.getMessage());
        }catch (Exception ex) {
            JFrame errorFrame = new JFrame(); // Create a new frame for the error dialog
            JOptionPane.showMessageDialog(
            	    errorFrame, 
            	    "An error occurred: Check if the column name is placed in the correct column data. If not, please solve and retry.", 
            	    "Error", 
            	    JOptionPane.ERROR_MESSAGE
            	);
//            System.out.println(ex.getMessage());
           
        }
    }

    private static String cleanValue(String value) {
        if (value == null || value.trim().isEmpty() || value.equals("null")) {
            return ""; // Return empty for unwanted values
        }
        return value;
    }

  

    private static String getCellValue(Cell cell) {
        if (cell == null) return null; // Return null for missing cells

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // Format date using DateTimeFormatter
                    LocalDate date = cell.getLocalDateTimeCellValue().toLocalDate();
                    return DateTimeFormatter.ofPattern("MM/dd/yyyy").format(date);
                } else {
                    // Handle numeric values (e.g., prices)
                    double value = cell.getNumericCellValue();
                    return formatPrice(value);
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                // Handle formulas based on the result type
                switch (cell.getCachedFormulaResultType()) {
                    case NUMERIC:
                        return String.valueOf(cell.getNumericCellValue());
                    case STRING:
                        return cell.getStringCellValue();
                    default:
                        return "";
                }
            default:
                return "";
        }
    }

    private static String formatPrice(double value) {
        DecimalFormat df = new DecimalFormat("#,###.00");
        return df.format(value);
    }

}






