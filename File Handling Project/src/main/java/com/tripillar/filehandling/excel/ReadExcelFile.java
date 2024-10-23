package com.tripillar.filehandling.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcelFile {

    public static void main(String[] args) {

        try {
            // Open the Excel file
            FileInputStream file = new FileInputStream(new File("StudentData.xlsx"));

            // Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get the first sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through all the rows in the sheet

            for (Row row : sheet) {
                // For each row, iterate through all the columns (cells)
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    // Check the cell type and print the appropriate value
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print((int) cell.getNumericCellValue() + "\t");
                            break;
                        default:
                            break;
                    }
                }
                System.out.println(); // Move to the next line after each row
            }

            // Close the file and workbook
            file.close();
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
