package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcel{
    private static final String FileDir = "/Users/soyrice/Desktop/javaL/WordProgram/src/excel/정리.xlsx";
    private static StringBuilder sb = new StringBuilder();

    public static StringBuilder excelGet(String targetWord) throws IOException {
        FileInputStream file = new FileInputStream(FileDir);
        IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        XSSFSheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                // 셀의 값이 null이 아니고, 시작하는 단어가 일치하는지 확인
                if (cell != null && cell.getStringCellValue().startsWith(targetWord) && !cell.getStringCellValue().contains("c.")) {
                    sb.append(cell.getStringCellValue());
                    sb.append(",");

                    int columnIndex = cell.getColumnIndex() - 1;
                    if (columnIndex >= 0) {
                        Cell leftCell = row.getCell(columnIndex);
                        if (leftCell != null) {

                            sb.append(leftCell.getStringCellValue());
                            sb.append(",");
                        }
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
