package exceldown.easyexceldownload.impl;

import exceldown.easyexceldownload.download.ExcelDownload;
import exceldown.easyexceldownload.exception.DownloadException;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DefaultExcelDownload implements ExcelDataSetting, ExcelDownload {

    private final List<HashMap<String,Object>> excelData;

    //헤더 이름
    String[] headerNames;

    //행 카운트
    public int rowCount = 0; // 데이터가 저장될 행

    // 엑셀 workbook 생성
    Workbook workbook = new XSSFWorkbook();

    int startRow = 2;
    HttpServletResponse res;

    public Sheet sheet = workbook.createSheet(); // 엑셀 sheet 이름

    @Override
    public void excelStyleSet() {
        // 디폴트 너비 설정
        sheet.setDefaultColumnWidth(28);

        /**
         * header font style
         */
        XSSFFont headerXSSFFont = (XSSFFont) workbook.createFont();
        headerXSSFFont.setColor(new XSSFColor(new byte[]{(byte) 255, (byte) 255, (byte) 255}));

        /**
         * header cell style
         */
        XSSFCellStyle headerXssfCellStyle = (XSSFCellStyle) workbook.createCellStyle();

        // 테두리 설정
        headerXssfCellStyle.setBorderLeft(BorderStyle.THIN);
        headerXssfCellStyle.setBorderRight(BorderStyle.THIN);
        headerXssfCellStyle.setBorderTop(BorderStyle.THIN);
        headerXssfCellStyle.setBorderBottom(BorderStyle.THIN);

        // 배경 설정
        headerXssfCellStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 34, (byte) 37, (byte) 41}));
        headerXssfCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerXssfCellStyle.setFont(headerXSSFFont);

        /**
         * body cell style
         */
        XSSFCellStyle bodyXssfCellStyle = (XSSFCellStyle) workbook.createCellStyle();

        // 테두리 설정
        bodyXssfCellStyle.setBorderLeft(BorderStyle.THIN);
        bodyXssfCellStyle.setBorderRight(BorderStyle.THIN);
        bodyXssfCellStyle.setBorderTop(BorderStyle.THIN);
        bodyXssfCellStyle.setBorderBottom(BorderStyle.THIN);

        /**
         * header data
         */
        Row headerRow = null;
        Cell headerCell = null;

        headerRow = sheet.createRow(rowCount++);
        for(int i=0; i< headerNames.length; i++) {
            headerCell = headerRow.createCell(i);
            headerCell.setCellValue(headerNames[i]); // 데이터 추가
            headerCell.setCellStyle(headerXssfCellStyle); // 스타일 추가
        }
    }

    public DefaultExcelDownload(HttpServletResponse res, List<HashMap<String ,Object>> calcs, String[] headerNames) throws IOException {
        this.excelData = calcs;
        this.res = res;
        this.headerNames = headerNames;
    }

    public DefaultExcelDownload(List<HashMap<String ,Object>> calcs, String[] headerNames) throws IOException {
        this.excelData = calcs;
        this.headerNames = headerNames;
    }

    @Override
    public void excelDataSet() {
        /**
         * body data
         */
        Row bodyRow = null;
        Cell bodyCell = null;

        Map<String, Object> poiSetting = new HashMap<>();
        poiSetting.put("bodyRow", bodyRow);
        poiSetting.put("bodyCell", bodyCell);
        poiSetting.put("rowCount", rowCount);

        for (HashMap<String, Object> excelOneData : excelData) {
            poiSetting.put("bodyRow", sheet.createRow((Integer) poiSetting.get("rowCount")));
            Row bodyRow_temp = (Row) poiSetting.get("bodyRow");
            Cell bodyCell_temp = (Cell) poiSetting.get("bodyCell");

            int columnIndex = 0; // 열 인덱스 초기화

            for (String key : excelOneData.keySet()) {

                Object value = excelOneData.get(key); // 현재 키에 대한 값 가져오기

                bodyCell_temp = bodyRow_temp.createCell(columnIndex); // 현재 열에 셀 생성

                // 셀에 값 설정
                if (value == null || value.toString().isEmpty()) {
                    bodyCell_temp.setCellValue("-"); // 값이 없는 경우 "-"로 설정
                } else {
                    bodyCell_temp.setCellValue(value.toString()); // 값이 있는 경우 해당 값으로 설정
                }

                columnIndex++; // 열 인덱스 증가
            }

            poiSetting.put("bodyRow", bodyRow_temp);
            poiSetting.put("bodyCell", bodyCell_temp);
            poiSetting.put("rowCount", (Integer) poiSetting.get("rowCount") + 1);
        };
    }
    @Override
    public void httpExcelDownloadSet(String fileName, HttpServletResponse serResponse) throws IOException {
        // 엑셀 스타일 설정
        excelStyleSet();

        // 엑셀 데이터 설정
        excelDataSet();

        /**
         * download
         */
        serResponse.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        serResponse.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");
        ServletOutputStream servletOutputStream = serResponse.getOutputStream();

        workbook.write(servletOutputStream);
        workbook.close();
        servletOutputStream.flush();
        servletOutputStream.close();
    }

    @Override
    public void localDownloadSet(String fileName, String filePath) throws DownloadException {
       // 엑셀 파일 경로 및 파일 이름 설정
       String file = filePath + "/" + fileName + ".xlsx";

        // 엑셀 스타일 설정
       excelStyleSet();

       // 엑셀 데이터 설정
       excelDataSet();

        // 파일로 저장
        try (FileOutputStream fileOut = new FileOutputStream(file)) {
            workbook.write(fileOut);
        } catch (IOException e) {
            throw new DownloadException("File Download Fail" , e);
        }
    }
}
