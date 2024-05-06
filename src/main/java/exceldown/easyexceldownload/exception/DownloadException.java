package exceldown.easyexceldownload.exception;

public class DownloadException extends ExcelException {

    public DownloadException(){
        super();
    }

    public DownloadException(String msg){
        super(msg);
    }

    public DownloadException(String msg, Exception e){
        super(msg, e);
    }

}
