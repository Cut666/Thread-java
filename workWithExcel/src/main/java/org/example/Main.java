package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType; // Thêm dòng này
import org.example.Control.CreateExcel;
import org.example.Control.ReadExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;
import java.io.FileOutputStream;

public class Main {
    public static void main(String[] args) {
//        ReadExcel excel = new ReadExcel();
//        excel.read();
//        CreateExcel excel1 = new CreateExcel();
//        excel1.create();
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
        Thread exportThread = new Thread(() -> {
            try (FileOutputStream fileOut = new FileOutputStream("D:/downloads/output.xlsx")) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            }
            System.out.println("Dữ liệu đã được nhập và lưu vào tệp Excel.");
        });
        Thread sumThread = new Thread(() -> {
            for (int row = 0; row < rows; row++) {
                int total = 0;
                Row sheetRow = sheet.getRow(row);
                for (int col = 0; col < cols; col++) {
                    total += Integer.parseInt(sheetRow.getCell(col).getStringCellValue());
                }
                sheetRow.createCell(cols).setCellValue(total);
            }
        });

        scanner.close();
        exportThread.start();
        sumThread.start();
        try {
            exportThread.join();
            sumThread.join();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }
}