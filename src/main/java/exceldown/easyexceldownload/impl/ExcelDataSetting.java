package exceldown.easyexceldownload.impl;

import exceldown.easyexceldownload.download.ExcelDownload;

public interface ExcelDataSetting extends ExcelDownload {
    //내용 세팅
    void excelDataSet();

    //엑셀 스타일 세팅
    void excelStyleSet();
}
