package org.example.Control;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.IOException;
import java.util.Scanner;
import java.io.FileOutputStream;
public class CreateExcel {
    public void create(){
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Danh sách");
        Scanner scanner = new Scanner(System.in);
        System.out.println("Nhập số dòng: ");
        int rows = scanner.nextInt();
        System.out.println("Nhập số cột: ");
        int cols = scanner.nextInt();

        for (int row = 0; row < rows; row++) {
            Row sheetRow = sheet.createRow(row);
            for (int col = 0; col < cols; col++) {
                System.out.println("Nhập giá trị cho ô [" + (row + 1) + "," + (col + 1) + "]: ");
                String cellValue = scanner.next();
                sheetRow.createCell(col).setCellValue(cellValue);
            }
        }

        System.out.println("Nhập đường dẫn cho tệp Excel (ví dụ: D:/download/output.xlsx): ");
        String filePath = scanner.next();

        scanner.close();
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Dữ liệu đã được nhập và lưu vào tệp Excel tại " + filePath);
    }

}
