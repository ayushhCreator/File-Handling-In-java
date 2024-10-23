package com.tripillar.filehandling.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class WriteExcelFile {

    public static void main(String[] args) {

        // Create a new workbook (Excel file)
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Create a blank sheet in the workbook named "Student Data"
        XSSFSheet sheet = workbook.createSheet("Student Data");

        // Prepare data to be written in the Excel file
        // TreeMap is used to store data in key-value pairs where each key corresponds to a row
        Map<String, Object[]> data = new TreeMap<>();
        // Adding header row and student data as Object arrays
        data.put("1", new Object[]{"ID", "NAME", "Course"});
        data.put("2", new Object[]{1, "Ayush", "MCA"});
        data.put("3", new Object[]{2, "Abhi", "MCA"});
        data.put("4", new Object[]{3, "Rajesh", "MBA"});
        data.put("5", new Object[]{4, "Suresh", "MBA"});
        data.put("6", new Object[]{5, "Naresh", "MCA"});

        // Iterate over data and write it to the sheet
        Set<String> keyset = data.keySet();  // Get the set of keys (row identifiers)
        int rownum = 0;  // Track row numbers
        for (String key : keyset) {
            // Create a new row in the Excel sheet for each key in the TreeMap
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);  // Get the data for each key (row)
            int cellnum = 0;  // Track cell numbers within each row
            for (Object obj : objArr) {
                // Create a new cell in the row and write data into it
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);  // Write String data
                } else if (obj instanceof Integer) {
                    cell.setCellValue((Integer) obj);  // Write Integer data
                }
            }
        }

        // Write the workbook to a file in the file system
        try {
            // Specify the file path where the Excel file will be saved
            FileOutputStream out = new FileOutputStream(new File("StudentData.xlsx"));
            workbook.write(out);  // Write the workbook data to the file
            out.close();  // Close the file output stream
            workbook.close();  // Close the workbook to prevent memory leaks
            System.out.println("StudentData.xlsx written successfully in the Excel Sheet.");  // Confirmation message
        } catch (Exception e) {
            e.printStackTrace();  // Print the stack trace in case of any exception
        }
    }
}
