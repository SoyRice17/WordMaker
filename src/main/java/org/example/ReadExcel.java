package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcel {
    private static final String FileDir = "/Users/soyrice/Desktop/javaL/WordProgram/src/excel/정리.xlsx";
    private static StringBuilder sb = new StringBuilder();

    public static StringBuilder excelGet(String targetWord, int targetColumnIndex) throws IOException {
        FileInputStream file = new FileInputStream(FileDir);
        IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        XSSFSheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            Cell cell = row.getCell(targetColumnIndex); // 특정 열에서 셀을 가져옴

            // 셀의 값이 null이 아니고, 시작하는 단어가 일치하며 "c."를 포함하지 않는 경우에만 값을 추가
            if (cell != null && cell.getStringCellValue().startsWith(targetWord) && !cell.getStringCellValue().contains("c.")) {
                sb.append(cell.getStringCellValue());
                sb.append(",");

                int leftColumnIndex = targetColumnIndex - 1;
                if (leftColumnIndex >= 0) {
                    Cell leftCell = row.getCell(leftColumnIndex);
                    if (leftCell != null) {
                        sb.append(leftCell.getStringCellValue());
                        sb.append(",");
                    }
                }
            }
        }
        file.close();
        return sb;
    }

    public static StringBuilder getSb() {
        return sb;
    }

    public static void setSb(StringBuilder sb) {
        ReadExcel.sb = sb;
    }
}
