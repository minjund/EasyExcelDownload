package exceldown.easyexceldownload;

import exceldown.easyexceldownload.download.ExcelDownload;
import exceldown.easyexceldownload.exception.ExcelException;
import exceldown.easyexceldownload.impl.DefaultExcelDownload;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * REST Controller for Excel download operations
 * Provides endpoints for generating and downloading Excel files
 */
@RestController
public class ExcelDownLoad {

    /**
     * HTTP endpoint for Excel file download
     * Generates a sample Excel file with test data
     */
    @GetMapping("/excelDownLoad")
    public void httpExcelDown(HttpServletResponse response) throws IOException {
        List<HashMap<String, Object>> data = createSampleData();
        String[] headers = new String[]{"Column 1", "Column 2", "Column 3"};

        ExcelDownload excelDownload = new DefaultExcelDownload(data, headers);
        excelDownload.httpExcelDownloadSet("httpFileDownload", response);
    }

    /**
     * Local file system Excel download
     * Saves Excel file to specified directory
     */
    public static void localExcelDown(List<HashMap<String, Object>> data, String[] headerNames,
                                     String fileName, String filePath) throws ExcelException {
        ExcelDownload excelDownload = new DefaultExcelDownload(data, headerNames);
        excelDownload.localDownloadSet(fileName, filePath);
    }

    /**
     * Creates sample data for demonstration purposes
     */
    private List<HashMap<String, Object>> createSampleData() {
        List<HashMap<String, Object>> data = new ArrayList<>();

        HashMap<String, Object> row1 = new HashMap<>();
        row1.put("0", "value1");
        row1.put("1", 123);
        data.add(row1);

        HashMap<String, Object> row2 = new HashMap<>();
        row2.put("0", "value2");
        row2.put("1", 456);
        data.add(row2);

        HashMap<String, Object> row3 = new HashMap<>();
        row3.put("0", "value3");
        row3.put("1", 567);
        row3.put("2", 789);
        data.add(row3);

        return data;
    }

    /**
     * Exception handler for Excel-related errors
     */
    @ExceptionHandler(ExcelException.class)
    public void handleExcelException(ExcelException ex, HttpServletResponse response) throws IOException {
        response.sendError(HttpStatus.INTERNAL_SERVER_ERROR.value(),
            "Excel generation failed: " + ex.getMessage());
    }

    /**
     * Exception handler for IO errors
     */
    @ExceptionHandler(IOException.class)
    public void handleIOException(IOException ex, HttpServletResponse response) throws IOException {
        response.sendError(HttpStatus.INTERNAL_SERVER_ERROR.value(),
            "File operation failed: " + ex.getMessage());
    }
}
