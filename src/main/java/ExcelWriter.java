import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.File;

public class ExcelWriter {
    public static void main(String[] args) {
        // Създаваме нов Excel файл
        Workbook workbook = new XSSFWorkbook(); // Използваме XSSFWorkbook за .xlsx формат

        // Създаваме нов лист
        Sheet sheet = workbook.createSheet("Products");

        // Създаваме редове и клетки с данни
        Row headerRow = sheet.createRow(0); // Първи ред за заглавията
        headerRow.createCell(0).setCellValue("Product");
        headerRow.createCell(1).setCellValue("Price");

        // Добавяме данни за продукти
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Apple");
        row1.createCell(1).setCellValue(1.25);

        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Banana");
        row2.createCell(1).setCellValue(0.85);

        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue("Orange");
        row3.createCell(1).setCellValue(1.45);

        // Проверка дали директорията съществува, ако не я създаваме
        File directory = new File("D:/Excel");
        if (!directory.exists()) {
            directory.mkdirs(); // Създава директорията, ако не съществува
        }

        // Пълен път до Excel файла
        String filePath = "D:/Excel/products.xlsx";
        
        // Записваме в Excel файла
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
            System.out.println("Excel файлът е създаден успешно!");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close(); // Затваряме workbook-а
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}