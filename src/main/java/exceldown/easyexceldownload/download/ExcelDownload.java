package exceldown.easyexceldownload.download;

import exceldown.easyexceldownload.exception.ExcelException;
import jakarta.servlet.http.HttpServletResponse;

import java.io.IOException;

public interface ExcelDownload {

    //엑셀 다운로드 방법 세팅
    void httpExcelDownloadSet(String fileName, HttpServletResponse servletResponse) throws ExcelException, IOException;

    void localDownloadSet(String fileName, String filePath) throws ExcelException;
}
