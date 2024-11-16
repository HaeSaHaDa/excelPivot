package solomonm.seocho.seochoPivot.main;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.stereotype.Component;

import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

@Component
public class SeochoPivotMain implements ApplicationRunner {

    @Override
    public void run(ApplicationArguments args) throws Exception {
        // CSV 파일 경로 설정
        String csvFilePath = "D:\\myProject\\SolomonPj\\seochogucheong\\solomonPivot\\tb_seocho_result.csv";
        String excelFilePath = "D:\\myProject\\SolomonPj\\seochogucheong\\solomonPivot\\converted_data.xlsx";
        System.out.println("=====================================================");
        // 엑셀 파일 생성 메서드 호출
        convertCsvToExcel(csvFilePath, excelFilePath);

    }
    public static void convertCsvToExcel(String csvFilePath, String excelFilePath) throws IOException {
        try (BufferedReader br = Files.newBufferedReader(Paths.get(csvFilePath));
             Workbook workbook = new SXSSFWorkbook()) {

            System.out.println(csvFilePath);
            System.out.println(excelFilePath);

            // 새 워크시트 생성
            Sheet sheet = workbook.createSheet("Sheet1");
            String line;
            int rowNum = 0;
            System.out.println("================[새 워크시트 생성]=================");

            while ((line = br.readLine()) != null) {
                String[] fields = line.split(",");
                Row row = sheet.createRow(rowNum++);

                for (int i = 0; i < fields.length; i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(fields[i]);
                }
            }
            System.out.println("================[로드 완료]=================");

            // 엑셀 파일 저장
            try (FileOutputStream fileOut = new FileOutputStream(excelFilePath);
                 BufferedOutputStream fileOutBuffer = new BufferedOutputStream(fileOut)) {
                workbook.write(fileOutBuffer);
            }

            System.out.println("================[저장 완료]=================");
            System.out.println("엑셀 파일이 성공적으로 생성되었습니다: " + excelFilePath);
        }
    }
}
