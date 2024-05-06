package exceldown.easyexceldownload;

import exceldown.easyexceldownload.impl.DefaultExcelDownload;
import exceldown.easyexceldownload.impl.ExcelDataSetting;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;


@RestController
public class ExcelDownLoad {
    @GetMapping("/excelDownLoad")
    public void httpExcelDown(HttpServletResponse res) throws IOException {
        List<HashMap<String, Object>> calcs = new ArrayList<HashMap<String, Object>>() {{
            add(new HashMap<String, Object>() {{
                put("0", "value1");
                put("1", 123);
            }});
            add(new HashMap<String, Object>() {{
                put("0", "value2");
                put("1", 456);
            }});
            add(new HashMap<String, Object>() {{
                put("0", "value3");
                put("1", 567);
                put("2", 567);
            }});
        }};

        String[] headerNames = new String[] {
                "1",
                "2",
                "3"
        };

        ExcelDataSetting excelDownload = new DefaultExcelDownload(res, calcs, headerNames);

        excelDownload.httpExcelDownloadSet("httpFileDownTest", res);
    }

    public static void localExcelDown(List<HashMap<String, Object>> calcs, String[] headerNames) throws IOException {
        //엑셀 다운 다형성
        ExcelDataSetting excelDownload = new DefaultExcelDownload(calcs, headerNames);

        excelDownload.localDownloadSet("localDownTest2","/Users/kimminjun/Downloads");

    }

    public static void main(String[] args) throws IOException {
        List<HashMap<String, Object>> calcs = new ArrayList<HashMap<String, Object>>() {{
            add(new HashMap<String, Object>() {{
                put("orderId", "value1");
                put("password", 123);
                put("password22", "");
            }});
            add(new HashMap<String, Object>() {{
                put("orderId", "value2");
                put("password", 456);
                put("password22", "ㅇㅅㅇㅅㅇ");
            }});
            add(new HashMap<String, Object>() {{
                put("orderId", "value3");
                put("password", 567);
                put("password22", "가나다라");
            }});
        }};

        String[] headerNames = new String[] {
                "1",
                "2",
                "3"
        };

        localExcelDown(calcs, headerNames);
    }
}
