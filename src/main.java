import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class main {

    private JFrame fileGui;
    private JFileChooser fileChooser;

    public static void main(String[] args) {
        EventQueue.invokeLater(() -> {
            try {
                main window = new main();
                window.fileGui.setVisible(true);
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
    }

    public main() {
        initialize();
    }

    private void initialize() {
        fileGui = new JFrame();
        fileGui.setTitle("Dateien umwandeln (CSV to Excel)");
        fileGui.setBounds(100, 100, 450, 300);
        fileGui.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        JPanel panel = new JPanel();
        fileGui.getContentPane().add(panel, BorderLayout.CENTER);
        JButton btnChangetoExcel = new JButton("CSV-Datei in Excel-Datei umwandeln");
        JButton btnCompareCSV = new JButton("Zusammenfügen von CSV-Dateien");
        JButton btnBothButtons = new JButton("Erst zusammenfügen, dann umwandeln.");
        JButton btnCompareAndSelectRows = new JButton("Zusammenfügen und Zeilen auswählen");
        panel.add(btnChangetoExcel);
        panel.add(btnCompareCSV);
        panel.add(btnBothButtons);
        fileGui.pack();
        fileGui.setLocationRelativeTo(null);
        fileChooser = new JFileChooser();
        fileChooser.setMultiSelectionEnabled(false);
        fileChooser.setCurrentDirectory(new File(System.getProperty("user.home") + "/Documents"));
        fileChooser.setFileFilter(new FileNameExtensionFilter("CSV files", "csv"));

        btnChangetoExcel.addActionListener(e -> convertToExcel());
        btnCompareCSV.addActionListener(e -> {
            fileChooser.setMultiSelectionEnabled(true);
            compareCSVFiles();
        });
        btnBothButtons.addActionListener(e -> {
            fileChooser.setMultiSelectionEnabled(true);
            compareBoth();
        });

        btnCompareAndSelectRows.addActionListener(e -> {
            fileChooser.setMultiSelectionEnabled(true);
            compareAndSelectRows();
        });


    }

    private void compareAndSelectRows() {

    }

    private void compareBoth() {
        int returnVal = fileChooser.showOpenDialog(fileGui);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File[] selectedFiles = fileChooser.getSelectedFiles();
            try {
                String fileName = selectedFiles[0].getName();
                String mergedFilePath = selectedFiles[0].getParent() + File.separator + fileName.replace(".csv", "-MERGED.csv");
                BufferedWriter writer = new BufferedWriter(new FileWriter(mergedFilePath));
                boolean isFirstFile = true;
                for (File selectedFile : selectedFiles) {
                    List<String> lines = Files.readAllLines(selectedFile.toPath());
                    if (isFirstFile) {
                        writer.write(lines.get(0));
                        isFirstFile = false;
                    }
                    for (int i = 1; i < lines.size(); i++) {
                        writer.newLine();
                        writer.write(lines.get(i));
                    }
                }
                writer.close();

                File csvFile = new File(mergedFilePath);
                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet("1");
                List<String> lines = Files.readAllLines(csvFile.toPath());
                int rownum = 0;
                for (String line : lines) {
                    XSSFRow row = sheet.createRow(rownum++);
                    String[] values = line.split(",");
                    int cellnum = 0;
                    for (String value : values) {
                        XSSFCell cell = row.createCell(cellnum++);
                        cell.setCellValue(value);
                    }
                }
                String xlsxFilePath = csvFile.getAbsolutePath().replace(".csv", ".xlsx");
                FileOutputStream out = new FileOutputStream(xlsxFilePath);
                workbook.write(out);
                out.close();
                workbook.close();

                JOptionPane.showMessageDialog(fileGui, "CSV-Dateien erfolgreich zusammengefügt und in eine Excel-Datei umgewandelt: " + xlsxFilePath);
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(fileGui, "Fehler!");
            }
        }
    }


    private void convertToExcel() {
        int returnVal = fileChooser.showOpenDialog(fileGui);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File csvFile = fileChooser.getSelectedFile();
            try {
                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet("1");
                List<String> lines = Files.readAllLines(csvFile.toPath());
                int rownum = 0;
                for (String line : lines) {
                    XSSFRow row = sheet.createRow(rownum++);
                    String[] values = line.split(",");
                    int cellnum = 0;
                    for (String value : values) {
                        XSSFCell cell = row.createCell(cellnum++);
                        cell.setCellValue(value);
                    }
                }
                String csvFilePath = csvFile.getAbsolutePath();
                Path xlsxPath = Paths.get(csvFilePath.substring(0, csvFilePath.lastIndexOf('.')) + "-FORMATTED.xlsx");
                FileOutputStream out = new FileOutputStream(xlsxPath.toFile());
                workbook.write(out);
                out.close();

                JOptionPane.showMessageDialog(fileGui, "Erfolgreich in eine formatierte Excel-Datei gespeichert!");
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(fileGui, "Fehler!");
            }
        }
    }

    private void compareCSVFiles() {
        int returnVal = fileChooser.showOpenDialog(fileGui);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File[] selectedFiles = fileChooser.getSelectedFiles();
            try {
                String fileName = selectedFiles[0].getName();
                String mergedFilePath = selectedFiles[0].getParent() + File.separator +
                        fileName.replace(".csv", "-MERGED.csv");
                BufferedWriter writer = new BufferedWriter(new FileWriter(mergedFilePath));
                boolean isFirstFile = true;
                for (File selectedFile : selectedFiles) {
                    List<String> lines = Files.readAllLines(selectedFile.toPath());
                    if (isFirstFile) {
                        writer.write(lines.get(0));
                        isFirstFile = false;
                    }
                    for (int i = 1; i < lines.size(); i++) {
                        writer.newLine();
                        writer.write(lines.get(i));
                    }
                }
                writer.close();
                JOptionPane.showMessageDialog(fileGui, "CSV Datei wurde erfolgreich zusammengefasst: " + mergedFilePath);
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(fileGui, "Fehler!");
            }
        }
    }
}
