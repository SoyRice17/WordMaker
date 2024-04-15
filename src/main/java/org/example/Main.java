package org.example;
import com.aspose.words.*;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;

public class Main {
    static BufferedReader br = new BufferedReader(new InputStreamReader(System.in));

    public static void main(String[] args) throws Exception {
        boolean running = true;
        while (running) {
            System.out.println("원하는 작업을 선택하세요:");
            System.out.println("1. 분류번호 읽어오기");
            System.out.println("2. 특수 분류번호 읽어오기");
            System.out.println("3. 종료");

            int choiceNum = Integer.parseInt(br.readLine());

            switch (choiceNum) {
                case 1:
                    processNormalNumbers();
                    break;
                case 2:
                    processSpecialNumbers();
                    break;
                case 3:
                    running = false;
                    System.out.println("프로그램을 종료합니다.");
                    break;
                default:
                    System.out.println("올바른 옵션을 선택하세요.");
            }
        }
    }

    private static void processNormalNumbers() throws Exception {
        String[] result = ReadExcel.excelGet(9);
        System.out.println("숫자가 들어있는 배열:");
        for (String element : result) {
            System.out.println(element);
        }

        // Excel 파일에서 데이터 가져오기
        String[] data = ReadExcel.excelGet(9);

        // 가져온 데이터를 ArrayList로 변환
        ArrayList<String> dataList = new ArrayList<>();
        for (String element : data) {
            String[] values = element.split(",");
            for (String value : values) {
                dataList.add(value.trim());
            }
        }

        // 데이터를 100개씩 묶어서 처리
        int batchSize = 100;
        int k = 0;
        while (k < dataList.size()) {
            // 문서 개체 만들기
            Document doc = new Document();
            // DocumentBuilder 개체 만들기
            DocumentBuilder builder = new DocumentBuilder(doc);
            // 표 생성
            Table table = builder.startTable();

            int endIndex = Math.min(k + batchSize, dataList.size());
            ArrayList<String> batchData = new ArrayList<>(dataList.subList(k, endIndex));

            // 표에 데이터 삽입
            for (int i = 0; i < batchData.size(); i += 10) {
                // 첫 번째 행 (공백으로 5열)
                for (int j = 0; j < 5; j++) {
                    builder.insertCell();
                    builder.write("");
                }
                builder.endRow();

                // 둘째 행 (1부터 홀수로 5열)
                for (int j = i + 1; j < i + 10 && j < batchData.size(); j += 2) {
                    builder.insertCell();
                    builder.write(batchData.get(j));
                }
                builder.endRow();

                // 셋째 행 (0부터 짝수로 5열)
                for (int j = i; j < i + 10 && j < batchData.size(); j += 2) {
                    builder.insertCell();
                    builder.write(batchData.get(j));
                }
                builder.endRow();
            }

            // 표 끝내기
            builder.endTable();
            // 문서 저장
            doc.save("NomalNumber" + (k / batchSize + 1) + ".docx");

            k += batchSize; // 다음 묶음을 처리하기 위해 인덱스를 증가시킴
        }
    }

    private static void processSpecialNumbers() throws Exception {
        String choiceT = br.readLine();
        String[] data = ReadExcel.excelGet2(choiceT, 9);
        String dataAsString = String.join(",", data);
        String[] newData = dataAsString.split(",");
        ArrayList<String> newDataList = new ArrayList<>();
        for (String element : newData) {
            if (!element.isEmpty()) {
                newDataList.add(element);
            }
        }

        // 문서 개체 만들기
        Document doc = new Document();
        // DocumentBuilder 개체 만들기
        DocumentBuilder builder = new DocumentBuilder(doc);
        // 테이블 생성
        Table table = builder.startTable();
        // 셀 삽입
        for (int i = 0; i < newDataList.size(); i += 10) {
            // 첫 번째 행 (공백으로 5열)
            for (int j = 0; j < 5; j++) {
                builder.insertCell();
                builder.write("");
            }
            builder.endRow();

            // 둘째 행 (1부터 홀수로 5열)
            for (int j = i + 1; j < i + 10 && j < newDataList.size(); j += 2) {
                builder.insertCell();
                builder.write(newDataList.get(j));
            }
            builder.endRow();

            // 셋째 행 (0부터 짝수로 5열)
            for (int j = i; j < i + 10 && j < newDataList.size(); j += 2) {
                builder.insertCell();
                builder.write(newDataList.get(j));
            }
            builder.endRow();


        }

        // 탁자
        builder.endTable();
        // 문서 저장
        doc.save("ChoiceText" + choiceT + ".docx");
    }
}
