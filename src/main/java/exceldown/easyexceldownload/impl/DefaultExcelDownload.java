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

/**
 * Default implementation of Excel download operations
 * Provides Excel file generation with styling and data population
 */
public class DefaultExcelDownload implements ExcelDownload {

    private final List<HashMap<String, Object>> excelData;
    private final String[] headerNames;
    private final Workbook workbook = new XSSFWorkbook();
    public final Sheet sheet = workbook.createSheet();

    private int currentRowIndex = 0;

    @Override
    public void setExcelStyle() {
        sheet.setDefaultColumnWidth(28);

        XSSFCellStyle headerStyle = createHeaderStyle();
        Row headerRow = sheet.createRow(currentRowIndex++);

        for (int columnIndex = 0; columnIndex < headerNames.length; columnIndex++) {
            Cell headerCell = headerRow.createCell(columnIndex);
            headerCell.setCellValue(headerNames[columnIndex]);
            headerCell.setCellStyle(headerStyle);
        }
    }

    /**
     * Creates the header cell style with white text on dark background
     */
    private XSSFCellStyle createHeaderStyle() {
        XSSFFont headerFont = (XSSFFont) workbook.createFont();
        headerFont.setColor(new XSSFColor(new byte[]{(byte) 255, (byte) 255, (byte) 255}));

        XSSFCellStyle headerStyle = (XSSFCellStyle) workbook.createCellStyle();
        setBorders(headerStyle);
        headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 34, (byte) 37, (byte) 41}));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setFont(headerFont);

        return headerStyle;
    }

    /**
     * Creates the body cell style with borders
     */
    private XSSFCellStyle createBodyStyle() {
        XSSFCellStyle bodyStyle = (XSSFCellStyle) workbook.createCellStyle();
        setBorders(bodyStyle);
        return bodyStyle;
    }

    /**
     * Sets thin borders on all sides of a cell style
     */
    private void setBorders(XSSFCellStyle style) {
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
    }

    /**
     * Constructor for creating Excel download handler
     *
     * @param excelData List of data rows to populate in Excel
     * @param headerNames Array of header column names
     */
    public DefaultExcelDownload(List<HashMap<String, Object>> excelData, String[] headerNames) {
        this.excelData = excelData;
        this.headerNames = headerNames;
    }

    @Override
    public void setExcelData() {
        for (HashMap<String, Object> rowData : excelData) {
            Row dataRow = sheet.createRow(currentRowIndex++);
            int columnIndex = 0;

            for (String key : rowData.keySet()) {
                Object value = rowData.get(key);
                Cell cell = dataRow.createCell(columnIndex++);

                if (value == null || value.toString().isEmpty()) {
                    cell.setCellValue("-");
                } else {
                    cell.setCellValue(value.toString());
                }
            }
        }
    }
    @Override
    public void httpExcelDownloadSet(String fileName, HttpServletResponse response) throws IOException {
        setExcelStyle();
        setExcelData();

        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");

        try (ServletOutputStream outputStream = response.getOutputStream()) {
            workbook.write(outputStream);
            workbook.close();
            outputStream.flush();
        }
    }

    @Override
    public void localDownloadSet(String fileName, String filePath) throws DownloadException {
        setExcelStyle();
        setExcelData();

        String fullPath = filePath + "/" + fileName + ".xlsx";

        try (FileOutputStream fileOut = new FileOutputStream(fullPath)) {
            workbook.write(fileOut);
            workbook.close();
        } catch (IOException e) {
            throw new DownloadException("File Download Fail", e);
        }
    }
}
