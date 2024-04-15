package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReadExcel {
    private static final String FileDir = "/Users/soyrice/Desktop/javaL/WordProgram/src/excel/정리.xlsx";
    private static StringBuilder sb = new StringBuilder();

    public static String[] excelGet(int targetColumnIndex) throws IOException {
        FileInputStream file = new FileInputStream(FileDir);
        IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        XSSFSheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();
        List<String> dataList = new ArrayList<>();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            Cell cell = row.getCell(targetColumnIndex); // 특정 열에서 셀을 가져옴

            // 셀의 값이 숫자로 시작하는 경우에만 값을 추가
            if (cell != null) {
                String cellValue = cell.getStringCellValue().trim();
                if (cellValue.matches("\\d+.*")&& !cell.getStringCellValue().contains("c.")) {
                    StringBuilder sb = new StringBuilder();
                    sb.append(cellValue).append(",");

                    int leftColumnIndex = targetColumnIndex - 1; // 옆 열의 인덱스
                    if (leftColumnIndex >= 0 && leftColumnIndex < row.getLastCellNum()) {
                        Cell leftCell = row.getCell(leftColumnIndex);
                        if (leftCell != null) {
                            sb.append(leftCell.getStringCellValue().trim()).append(",");
                        }
                    }

                    dataList.add(sb.toString());
                }
            }
        }

        file.close();

        // 데이터를 파싱하여 반환할 배열 생성
        String[] dataArrays = new String[dataList.size()];
        for (int i = 0; i < dataList.size(); i++) {
            String data = dataList.get(i);
            dataArrays[i] = data;
        }
        return dataArrays;
    }





    public static String[] excelGet2(String targetWord, int targetColumnIndex) throws IOException {
        FileInputStream file = new FileInputStream(FileDir);
        IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        XSSFSheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();
        List<String> dataList = new ArrayList<>();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            Cell cell = row.getCell(targetColumnIndex); // 특정 열에서 셀을 가져옴

            // 셀의 값이 null이 아니고, 시작하는 단어가 일치하며 "c."를 포함하지 않는 경우에만 값을 추가
            if (cell != null && cell.getStringCellValue().startsWith(targetWord) && !cell.getStringCellValue().contains("c.")) {
                StringBuilder sb = new StringBuilder();
                sb.append(cell.getStringCellValue()).append(",");

                int leftColumnIndex = targetColumnIndex - 1;
                if (leftColumnIndex >= 0) {
                    Cell leftCell = row.getCell(leftColumnIndex);
                    if (leftCell != null) {
                        sb.append(leftCell.getStringCellValue()).append(",");
                    }
                }
                dataList.add(sb.toString());
            }
        }
        file.close();

        // 데이터를 파싱하여 반환할 배열 생성
        String[] dataArrays = new String[dataList.size()];
        for (int i = 0; i < dataList.size(); i++) {
            String data = dataList.get(i);
            String[] tokens = data.split(",");
            StringBuilder processedData = new StringBuilder();
            for (int j = 0; j < tokens.length; j++) {
                if (j % 2 == 1 && tokens[j].length() > 5) {
                    tokens[j] = tokens[j].substring(5);
                }
                processedData.append(tokens[j]).append(",");
            }
            dataArrays[i] = processedData.toString();
        }
        return dataArrays;
    }

    public static StringBuilder getSb() {
        return sb;
    }

    public static void setSb(StringBuilder sb) {
        ReadExcel.sb = sb;
    }
}
