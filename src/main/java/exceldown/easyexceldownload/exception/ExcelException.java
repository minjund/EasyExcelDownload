package exceldown.easyexceldownload.exception;

public class ExcelException extends RuntimeException {

    protected ExcelException() {
        super();
    }

    protected ExcelException(String msg) {
        super(msg);
    }
    public ExcelException(String msg, Exception e) {
        super(msg, e);
    }


}
