package org.example;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.StringTokenizer;

public class Main {
    static BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
    static StringBuilder sb = new StringBuilder();
    public static void main(String[] args) throws IOException {

        System.out.println("처리할 작업 입력 : ");
        System.out.println("1. 분류번호 읽어오기");
        System.out.println("2. 특수 분류번호 읽어오기");
        int coiceNum = Integer.parseInt(br.readLine());

        if (coiceNum == 1) {

        }
        if (coiceNum == 2) {
            String data = Arrays.toString(ReadExcel.excelGet2(br.readLine(), 9));
            System.out.println(data);
        }

    }
}