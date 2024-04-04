package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.Iterator;

public class ExelCrwal {
    public static void main(String[] args) {
        try (FileInputStream file = new FileInputStream("/Users/soyrice/Desktop/javaL/WordProgram/src/excel/image.xlsx")) {
            IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE); // Consider removing this line
            BufferedReader br = new BufferedReader(new InputStreamReader(System.in));

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            int columnIndex = Integer.parseInt(br.readLine()); // 예시로 3번째 열을 선택

            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                Cell cell = row.getCell(columnIndex); // 선택한 열의 셀을 가져옴

                // 셀이 null이 아닌 경우에만 처리
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        default:
                            throw new IllegalStateException("Unexpected value: " + cell.getCellType());
                    }
                }
                System.out.println(""); // 각 행의 처리가 끝난 후 개행
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}