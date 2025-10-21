package exceldown.easyexceldownload.download;

import exceldown.easyexceldownload.exception.ExcelException;
import jakarta.servlet.http.HttpServletResponse;

import java.io.IOException;

/**
 * Interface for Excel file download operations
 * Supports both HTTP-based downloads and local file system storage
 */
public interface ExcelDownload {

    /**
     * Sets up data for the Excel file
     */
    void setExcelData();

    /**
     * Sets up styling for the Excel file
     */
    void setExcelStyle();

    /**
     * Downloads Excel file via HTTP response
     *
     * @param fileName The name of the file (without extension)
     * @param servletResponse The HTTP servlet response
     * @throws ExcelException If Excel generation fails
     * @throws IOException If file I/O fails
     */
    void httpExcelDownloadSet(String fileName, HttpServletResponse servletResponse) throws ExcelException, IOException;

    /**
     * Saves Excel file to local file system
     *
     * @param fileName The name of the file (without extension)
     * @param filePath The directory path to save the file
     * @throws ExcelException If Excel generation fails
     */
    void localDownloadSet(String fileName, String filePath) throws ExcelException;
}
