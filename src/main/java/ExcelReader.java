import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelReader {
    public static void main(String[] args) {

        String filePath = "D:/Excel/products.xlsx"; // Път към входния файл
        String outputFilePath = "D:/Excel/filtered_products.xlsx"; // Път към новия файл с филтрирани данни

        try (FileInputStream fis = new FileInputStream(filePath)) {
            // Зареждаме Excel файла като XSSFWorkbook
            Workbook workbook = new XSSFWorkbook(fis);

            // Избиране на първия лист от Excel файла
            Sheet sheet = workbook.getSheetAt(0);

            // Записване на обработените данни в нов Excel файл
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("FilteredData");

            double totalPrice = 0;
            int count = 0;
            int rowIndex = 0;

            // Обхождаме редовете на листа
            for (Row row : sheet) {
                // Пропускаме първия ред (заглавия)
                if (row.getRowNum() == 0) continue;

                // Пример: предполагаме, че цената е във втората колона
                Cell priceCell = row.getCell(1);
                if (priceCell != null && priceCell.getCellType() == CellType.NUMERIC) {
                    double price = priceCell.getNumericCellValue();
                    if (price > 1.0) { // Филтрираме редовете с цена > 1.0


                        // Записваме реда в новия файл
                        Row newRow = outputSheet.createRow(rowIndex++);
                        for (Cell cell : row) {
                            Cell newCell = newRow.createCell(cell.getColumnIndex());
                            switch (cell.getCellType()) {
                                case STRING:
                                    newCell.setCellValue(cell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    newCell.setCellValue(cell.getNumericCellValue());
                                    break;
                                default:
                                    newCell.setCellValue("Invalid Cell Type");
                            }
                        }
                        // Добавяме стойността на цената
                        totalPrice += price;
                        count++;
                    }
                }
            }

            // Изчисляваме средната стойност
            double averagePrice = count > 0 ? totalPrice / count : 0;
            System.out.println("Средната стойност на цените: " + averagePrice);

            // Записваме новия файл
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                outputWorkbook.write(fos);
                System.out.println("Excel файлът с филтрирани данни е създаден успешно!");
            }

            workbook.close();
            outputWorkbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}