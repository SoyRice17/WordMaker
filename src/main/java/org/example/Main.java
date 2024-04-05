package org.example;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.StringTokenizer;

public class Main {
    static BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
    public static void main(String[] args) throws IOException {

        String excelData = String.valueOf(ReadExcel.excelGet(br.readLine(), 9));
        StringTokenizer st = new StringTokenizer(excelData, ",");

        String[] dataArrays = new String[st.countTokens()];
        for (int i = 0; i <= st.countTokens(); i++){
            dataArrays[i] = st.nextToken();
            if (i %2 == 1){
                if (dataArrays[i].length() > 5) { // 문자열이 5문자 이상일 때만 처리
                    dataArrays[i] = dataArrays[i].substring(5);
                } else {
                    dataArrays[i] = ""; // 5문자 미만인 경우 빈 문자열로 처리
                }
            }
            System.out.println(dataArrays[i]);
        }


    }
}