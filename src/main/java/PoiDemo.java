import java.awt.Desktop;
import java.awt.GridLayout;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PoiDemo {
    public static void main(String[] args) {
        JFrame frame = new JFrame("Zajuns ILY");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(600, 300);
        frame.setLocationRelativeTo(null);
        frame.setLayout(new GridLayout(3, 1, 10, 10));

        JPanel filePanel = new JPanel();
        JTextField filePathField = new JTextField(20);
        filePathField.setEditable(false);
        JButton chooseFileButton = new JButton("Выбрать файл");
        filePanel.add(filePathField);
        filePanel.add(chooseFileButton);

        JPanel valuePanel = new JPanel();
        JTextField valueField = new JTextField(10);
        valuePanel.add(new JLabel("Введите новое значение: "));
        valuePanel.add(valueField);

        JPanel processPanel = new JPanel();
        JButton processButton = new JButton("Обработать");
        processPanel.add(processButton);

        frame.add(filePanel);
        frame.add(valuePanel);
        frame.add(processPanel);

        chooseFileButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
                public boolean accept(File f) {
                    return f.getName().toLowerCase().endsWith(".xls") || 
                           f.getName().toLowerCase().endsWith(".xlsx") ||
                           f.isDirectory();
                }
                public String getDescription() {
                    return "Excel Files (*.xls, *.xlsx)";
                }
            });
            
            int result = fileChooser.showOpenDialog(frame);
            if (result == JFileChooser.APPROVE_OPTION) {
                filePathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
            }
        });

        processButton.addActionListener(e -> {
            String filePath = filePathField.getText();
            if (filePath.isEmpty()) {
                JOptionPane.showMessageDialog(frame, "Пожалуйста, выберите файл!");
                return;
            }

            try {
                double newValue = Double.parseDouble(valueField.getText());
                processExcelFile(filePath, newValue);
            } catch (NumberFormatException ex) {
                JOptionPane.showMessageDialog(frame, "Пожалуйста, введите корректное числовое значение!");
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(frame, "Ошибка при обработке файла: " + ex.getMessage());
            }
        });

        frame.setVisible(true);
    }

    private static void processExcelFile(String filePath, double newValueColumn8) {
        try (InputStream inputStream = new FileInputStream(filePath)) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            int lastRowIndex = findLastRowIndexBeforeTagastus(sheet, 8, "tagastus").get(1);
            Cell N = sheet.getRow(sheet.getPhysicalNumberOfRows()).getCell(1);
            Cell V = sheet.getRow(sheet.getPhysicalNumberOfRows()+1).getCell(1);
            double tagastus = Math.round(sheet.getRow(findLastRowIndexBeforeTagastus(sheet,1,"Müügid (Arve-Saateleht)").get(0)-1).getCell(8).getNumericCellValue() * 100.0) / 100.0;

            if (lastRowIndex != -1) {
                newValueColumn8 = Math.round(newValueColumn8 * 100.0) / 100.0;
                updateCellValue(sheet, lastRowIndex, 8, newValueColumn8);

                double adjacentCellValue = getCellValue(sheet, lastRowIndex, 9);
                double updatedValueColumn7 = Math.round((newValueColumn8 + adjacentCellValue) * 100.0) / 100.0;
                updateCellValue(sheet, lastRowIndex, 7, updatedValueColumn7);
                
                double difference = Math.round((newValueColumn8 - tagastus) * 100.0) / 100.0;
                String stringValue = String.format("%.2f", difference).replace(",", ".");
                N.setCellValue(stringValue);
                V.setCellValue(stringValue);
                
                double updatedValueColumn6 = Math.round((updatedValueColumn7 * 22.0 / 122.0) * 100.0) / 100.0;
                updateCellValue(sheet, lastRowIndex, 6, updatedValueColumn6);

                double updatedValueColumn4 = Math.round((updatedValueColumn7 - updatedValueColumn6) * 100.0) / 100.0;
                updateCellValue(sheet, lastRowIndex, 4, updatedValueColumn4);

                if (newValueColumn8 == 0) {
                    deleteCell(sheet, N);
                    deleteCell(sheet, V);
                    deleteCell(sheet, lastRowIndex, 8);
                }
            } else {
                JOptionPane.showMessageDialog(null, "Строка 'tagastus' не найдена. Не удалось обновить значения в столбцах 4, 6, 7 и 8.");
                return;
            }

            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
                JOptionPane.showMessageDialog(null, "Значения успешно обновлены в файле Excel.");
                
                try {
                    Desktop.getDesktop().open(new File(filePath));
                } catch (Exception e) {
                    JOptionPane.showMessageDialog(null, "Не удалось открыть файл: " + e.getMessage());
                }
            }

            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Ошибка при обработке файла: " + e.getMessage());
        }
    }

    private static List<Integer> findLastRowIndexBeforeTagastus(Sheet sheet, int columnIndex, String name) {
        int lastRowIndex = -1;
        int newIndex = 0;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null && cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().equalsIgnoreCase(name)) {
                    newIndex = i;
                    break;
                }
                cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    lastRowIndex = i;
                }
            }
        }
        return Arrays.asList(newIndex, lastRowIndex);
    }

    private static double getCellValue(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                return Math.round(cell.getNumericCellValue() * 100.0) / 100.0;
            }
        }
        return 0.0;
    }

    private static void updateCellValue(Sheet sheet, int rowIndex, int columnIndex, double newValue) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex, CellType.NUMERIC);
        }
        cell.setCellValue(Math.round(newValue * 100.0) / 100.0);
    }

    private static void deleteCell(Sheet sheet, Cell cell) {
        Row row = cell.getRow();
        if (row != null) {
            row.removeCell(cell);
        }
    }

    private static void deleteCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null) {
                row.removeCell(cell);
            }
        }
    }
}
