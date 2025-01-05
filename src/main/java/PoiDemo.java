import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PoiDemo {
    public static void main(String[] args) throws IOException {
        String filePath = "C:\\Users\\dkart\\OneDrive\\Рабочий стол\\Учебные материалы\\Pp2.xls";

        try (InputStream inputStream = new FileInputStream(filePath)) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0); // Получаем первый лист книги

            // Получаем новое значение для столбца 8 с клавиатуры
            Scanner scanner = new Scanner(System.in);
            System.out.print("Введите новое значение для столбца 8: ");
            double newValueColumn8 = scanner.nextDouble();
            // Находим последнее число в столбце 8 перед строкой "tagastus" в столбце 0
            int lastRowIndex = findLastRowIndexBeforeTagastus(sheet, 8,"tagastus").get(1);
            Cell N = sheet.getRow(sheet.getPhysicalNumberOfRows()).getCell(1);
            Cell V = sheet.getRow(sheet.getPhysicalNumberOfRows()+1).getCell(1);
            double tagastus = sheet.getRow(findLastRowIndexBeforeTagastus(sheet,1,"Müügid (Arve-Saateleht)").get(0)-1).getCell(8).getNumericCellValue();
            // Обновляем значение только в последней ячейке столбца 8.
            if (lastRowIndex != -1) {
                // Обновляем значение только в последней ячейке столбца 8
                updateCellValue(sheet, lastRowIndex, 8, newValueColumn8);

                // Получаем значение ячейки справа от последней ячейки столбца 8
                double adjacentCellValue = getCellValue(sheet, lastRowIndex, 9);
                // Вычисляем новое значение как сумму значения ячейки слева и значения ячейки справа от этой ячейки
                double updatedValueColumn7 = newValueColumn8 + adjacentCellValue;
                // Обновляем значение только в последней ячейке столбца 7
                updateCellValue(sheet, lastRowIndex, 7, updatedValueColumn7);
                String stringValue = Double.toString(newValueColumn8 - tagastus); // Convert the result to a string
                N.setCellValue(stringValue+"0");
                V.setCellValue(stringValue+"0");
                // Вычисляем новое значение для столбца 6
                double updatedValueColumn6 = Math.round((updatedValueColumn7 * 24 / 124) * 100.0) / 100.0;
                // Обновляем значение только в последней ячейке столбца 6
                updateCellValue(sheet, lastRowIndex, 6, updatedValueColumn6);

                // Вычисляем новое значение для столбца 4
                double updatedValueColumn4 = Math.round((updatedValueColumn7 - updatedValueColumn6) * 100.0) / 100.0;
                // Обновляем значение только в последней ячейке столбца 4
                updateCellValue(sheet, lastRowIndex, 4, updatedValueColumn4);

                // If N is zero, delete cells N, V, and newValueColumn8
                if (newValueColumn8 == 0) {
                    deleteCell(sheet, N);
                    deleteCell(sheet, V);
                    deleteCell(sheet, lastRowIndex, 8); // Deleting cell at lastRowIndex and column 8
                }
            } else {
                System.out.println("Строка 'tagastus' не найдена. Не удалось обновить значения в столбцах 4, 6, 7 и 8.");
            }

            // Записываем изменения обратно в файл Excel
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            } catch (IOException e) {
                e.printStackTrace();
            }

            System.out.println("Значения успешно обновлены в файле Excel.");

            workbook.close(); // Закрываем книгу
        } catch (IOException | EncryptedDocumentException e) {
            e.printStackTrace();
        }
    }

    // Метод для поиска индекса последней строки перед строкой "tagastus" в указанном столбце
    private static List<Integer> findLastRowIndexBeforeTagastus(Sheet sheet, int columnIndex, String name) {

        int lastRowIndex = -1;
        int newIndex = 0;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null && cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().equalsIgnoreCase(name)) {
                    newIndex = i;
                    break; // Нашли строку "tagastus", выходим из цикла
                }
                cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    lastRowIndex = i; // Запоминаем индекс последней строки с числовым значением
                }
            }
        }
        return Arrays.asList(newIndex,lastRowIndex);

    }
    private static List<Integer> findSissemaks(Sheet sheet, int columnIndex, String name) {

        int lastRowIndex = -1;
        int newIndex = 0;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null && cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().equalsIgnoreCase(name)) {
                    newIndex = i;
                    break; // Нашли строку "tagastus", выходим из цикла
                }
                cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    lastRowIndex = i; // Запоминаем индекс последней строки с числовым значением
                }
            }
        }
        return Arrays.asList(newIndex,lastRowIndex);

    }

    // Метод

    // Метод для получения значения ячейки в указанной строке и столбце
    private static double getCellValue(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                return cell.getNumericCellValue();
            }
        }
        return 0.0; // Если значение не найдено или не числовое, возвращаем 0.0
    }

    // Метод для обновления значения ячейки в указанной строке и столбце
    private static void updateCellValue(Sheet sheet, int rowIndex, int columnIndex, double newValue) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex, CellType.NUMERIC);
        }
        cell.setCellValue(newValue);
    }

    // Method to delete a cell
    private static void deleteCell(Sheet sheet, Cell cell) {
        Row row = cell.getRow();
        if (row != null) {
            row.removeCell(cell);
        }
    }

    // Method to delete a cell at specified row and column index
    private static void deleteCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null) {
                row.removeCell(cell);
            }
        }
    }

    // Добавим новый вспомогательный метод для округления
    private static double roundToTwoDecimals(double value) {
        return Math.round(value * 100.0) / 100.0;
    }
}
