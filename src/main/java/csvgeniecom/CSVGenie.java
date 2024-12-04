package csvgeniecom;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

class XLSXRecord {
    private String code;
    private String name;
    private String type;
    private String taxCode;
    private String description;
    private String dashboard;
    private String expenseClaims;
    private String enablePayments;
    private String balance;

    // Getters and setters (same as in your original code)
    public String getCode() { return code; }
    public void setCode(String code) { this.code = code; }
    public String getName() { return name; }
    public void setName(String name) { this.name = name; }
    public String getType() { return type; }
    public void setType(String type) { this.type = type; }
    public String getTaxCode() { return taxCode; }
    public void setTaxCode(String taxCode) { this.taxCode = taxCode; }
    public String getDescription() { return description; }
    public void setDescription(String description) { this.description = description; }
    public String getDashboard() { return dashboard; }
    public void setDashboard(String dashboard) { this.dashboard = dashboard; }
    public String getExpenseClaims() { return expenseClaims; }
    public void setExpenseClaims(String expenseClaims) { this.expenseClaims = expenseClaims; }
    public String getEnablePayments() { return enablePayments; }
    public void setEnablePayments(String enablePayments) { this.enablePayments = enablePayments; }
    public String getBalance() { return balance; }
    public void setBalance(String balance) { this.balance = balance; }
}

public class CSVGenie {

    public static void main(String[] args) {
        SwingUtilities.invokeLater(CSVGenie::createAndShowGUI);
    }

    private static void createAndShowGUI() {
        JFrame frame = new JFrame("CSVGenie");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(500, 400);
        frame.setResizable(false);

        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BorderLayout());
        mainPanel.setBackground(Color.WHITE);
        mainPanel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));

        JLabel headerLabel = new JLabel("CSVGenie", JLabel.CENTER);
        headerLabel.setFont(new Font("Arial", Font.BOLD, 24));
        headerLabel.setForeground(new Color(50, 120, 200));

        JPanel centerPanel = new JPanel();
        centerPanel.setLayout(new GridBagLayout());
        centerPanel.setBackground(Color.WHITE);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(10, 10, 10, 10);

        JLabel instructions = new JLabel("1. Select multiple XLSX files.\n 2. Choose the output directory.");
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

        // File selection for input and output
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setMultiSelectionEnabled(true);
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));

        JFileChooser outputChooser = new JFileChooser();
        outputChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        List<File> selectedFiles = new ArrayList<>();
        File[] outputDirectory = new File[1]; // Store chosen output directory

        chooseFilesButton.addActionListener(e -> {
            int returnValue = fileChooser.showOpenDialog(frame);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                selectedFiles.clear();
                for (File file : fileChooser.getSelectedFiles()) {
                    selectedFiles.add(file);
                }
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

            // Create a modal dialog for the progress
            JDialog progressDialog = new JDialog(frame, "Processing...", true);
            progressDialog.setSize(300, 100);
            progressDialog.setLayout(new BorderLayout());
            progressDialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
            JLabel progressLabel = new JLabel("Converting files, please wait...", JLabel.CENTER);
            progressDialog.add(progressLabel, BorderLayout.CENTER);
            progressDialog.setLocationRelativeTo(frame);

            // SwingWorker for background processing
            SwingWorker<Void, Void> worker = new SwingWorker<>() {
                @Override
                protected Void doInBackground() {
                    try {
                        for (File file : selectedFiles) {
                            String outputFilePath = new File(outputDirectory[0], file.getName().replace(".xlsx", ".csv")).getAbsolutePath();
                            List<XLSXRecord> records = readXLSXFile(file.getAbsolutePath());
                            writeToCSV(records, outputFilePath);
                        }
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(frame, "An error occurred: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                    }
                    return null;
                }

                @Override
                protected void done() {
                    try {
                        // Add a delay to allow the user to see the completion status
                        Thread.sleep(2000); // 2 seconds delay
                    } catch (InterruptedException ignored) {
                        // Do nothing if interrupted
                    }
                    progressDialog.dispose(); // Close the progress dialog
                    statusLabel.setText("Conversion complete! Files saved in: " + outputDirectory[0].getAbsolutePath());
                }
            };

            worker.execute(); // Start the background task
            progressDialog.setVisible(true); // Show the progress dialog
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
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Read the first sheet
            int rowCount = 0; // Initialize row count to track rows

            for (Row row : sheet) {
                // Skip the header row and the second row
                if (row.getRowNum() == 0 || row.getRowNum() == 1) {
                    continue;
                }

                String nameValue = getCellValue(row.getCell(2)); // Name (Column C)
                
                // Check if the Name matches the keyword "Total" or "Total Toll"
                if (nameValue != null && nameValue.matches("(?i).*total.*")) {
                    // Skip the row immediately before the "Total" row
                    if (rowCount > 1) {
                        records.remove(records.size() - 1); // Remove the previous row
                    }
                    break; // Stop processing if regex matches
                }

                XLSXRecord record = new XLSXRecord();

                // Read basic details
                record.setCode(getCellValue(row.getCell(1))); // Code (Column B)
                record.setName(nameValue);

                // Read Debit and Credit values
                String debitValue = getCellValue(row.getCell(3)); // Debit (Column D)
                String creditValue = getCellValue(row.getCell(4)); // Credit (Column E)

                // Parse Debit and Credit to numbers
                double debit = parseNumericValue(debitValue);
                double credit = parseNumericValue(creditValue);

                double balance = debit - credit;
                record.setBalance(String.valueOf(balance));


                // Add the record to the list
                records.add(record);

                rowCount++; // Increment row count
            }
        }
        return records;
    }

 // Utility method to parse numeric values from string
    private static double parseNumericValue(String value) {
        try {
            return Double.parseDouble(value.trim());
        } catch (NumberFormatException e) {
            return 0; // Default to 0 if parsing fails
        }
    }
    public static void writeToCSV(List<XLSXRecord> records, String filePath) throws IOException {
        try (BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(filePath), "UTF-8"))) {
            // Write the header row
            String header = "*Code,*Name,*Type,*Tax Code,Description,Dashboard,Expense Claims,Enable Payments,Balance";
            bw.write(header);
            bw.newLine();

            // Write each record
            for (XLSXRecord record : records) {
                String row = String.join(",",
                        "\"" + cleanValue(record.getCode()) + "\"",
                        "\"" + cleanValue(record.getName()) + "\"",
                        "\"" + cleanValue(record.getType()) + "\"",
                        "\"Bas Excluded\"",
                        "\"" + cleanValue(record.getDescription()) + "\"",
                        "\"No\"",
                        "\"No\"",
                        "\"No\"",
                        "\"" + cleanValue(record.getBalance()) + "\"");

                bw.write(row);
                bw.newLine();
            }
        }
    }

    private static String cleanValue(String value) {
        if (value == null || value.trim().isEmpty() || value.equals("null") || value.contains("ï¾„")) {
            return ""; // Return empty for unwanted values
        }
        return value;
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return ""; // Return empty string for null cells

        String value = "";
        switch (cell.getCellType()) {
            case STRING:
                value = cell.getStringCellValue().trim();
                break;
            case NUMERIC:
                value = String.valueOf(cell.getNumericCellValue());
                break;
            case BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                try {
                    value = String.valueOf(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    value = cell.getStringCellValue();
                }
                break;
            default:
                return "";
        }

        // Return empty string if value is null, empty or contains unwanted characters
        if (value == null || value.trim().isEmpty() || value.equals("null") || value.contains("ï¾„")) {
            return "";
        }

        return value;
    }

}
