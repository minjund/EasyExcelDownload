package exceldown.easyexceldownload.excelMerge;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class SellerExcelMerge {

    private CellRangeAddress payment = new CellRangeAddress(0,0,0,0);
    private CellRangeAddress leedsDiscount = new CellRangeAddress(0,0,0,0);
    private CellRangeAddress realPayment = new CellRangeAddress(0,0,0,0);
    private CellRangeAddress deliveryMerge = new CellRangeAddress(0,0,0,0);

    private void defaultSetting(int[] formulaCellNum, int startRowIndex, int endRowIndex) {
        payment = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[0], formulaCellNum[0]);
        leedsDiscount = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[1], formulaCellNum[1]);
        realPayment = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[2], formulaCellNum[2]);
        deliveryMerge = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[3], formulaCellNum[3]);
    }

    public void sellerMergeValue(Sheet sheet, int totalNumberOfRows, int[] mergeCellNum, int[] formulaCellNum) {
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        List<String> formulaList = new ArrayList<>();
        int startRowIndex = 0;
        int endRowIndex = 0;

        for (int i = 0; i <= totalNumberOfRows; i++) {
            Row currentRow = sheet.getRow(i);
            if (currentRow == null) continue; // 현재 행이 없으면 건너뜁니다.

            //주문번호 value
            String orderIdValue = currentRow.getCell(0).getStringCellValue();
            String prevOrderIdValue = i > 0 ? sheet.getRow(i - 1).getCell(0).getStringCellValue() : null;

            if(orderIdValue.equals(prevOrderIdValue)) {
                endRowIndex = i;
            } else {
                if(startRowIndex < endRowIndex) {

                    for (int k : mergeCellNum) {
                        CellRangeAddress range = new CellRangeAddress(startRowIndex, endRowIndex, k, k);
                        setSellerFormula(sheet, formulaCellNum, k, startRowIndex, endRowIndex, formulaList);
                        mergedRegions.add(range);
                    }
                }

                if(startRowIndex > endRowIndex) {
                    for (int k : mergeCellNum) {
                        CellRangeAddress range = new CellRangeAddress(startRowIndex, startRowIndex, k, k);
                        setSellerFormula(sheet, formulaCellNum, k, startRowIndex, startRowIndex, formulaList);
                        mergedRegions.add(range);
                    }
                }

                startRowIndex = i;
            }
        }

        for (int i : mergeCellNum) {
            if (startRowIndex < endRowIndex) {
                CellRangeAddress range = new CellRangeAddress(startRowIndex, endRowIndex, i, i);
                setSellerFormula(sheet, formulaCellNum, i, startRowIndex, endRowIndex, formulaList);
                mergedRegions.add(range);
            } else {
                CellRangeAddress range = new CellRangeAddress(startRowIndex, startRowIndex, i, i);
                setSellerFormula(sheet, formulaCellNum, i, startRowIndex, startRowIndex, formulaList);
                mergedRegions.add(range);
            }
        }


        setMergedCell(sheet, mergedRegions, formulaList);
    }



    private void setSellerFormula(Sheet sheet, int[] formulaCellNum, int k, int startRowIndex, int endRowIndex, List<String> formulaList) {

        defaultSetting(formulaCellNum, startRowIndex, endRowIndex);

        String deliveryText = deliveryMerge.formatAsString().split(":")[0];

        generateFormula(sheet, k, startRowIndex, formulaList, deliveryText);
    }

    private void generateFormula(Sheet sheet, int k, int startRowIndex, List<String> formulaList, String deliveryText) {
        if(k == 11){
            formulaList.add("SUM("+ realPayment.formatAsString()+")");
        } else if (k == 12) {
            String leedsCouponAmount  = sheet.getRow(startRowIndex).getCell(k).getStringCellValue();
            sheet.getRow(startRowIndex).getCell(k).setCellValue(Integer.parseInt(leedsCouponAmount));
            sheet.getRow(startRowIndex).getCell(k).setCellType(CellType.NUMERIC);
        } else if (k == 13) {
            formulaList.add("(SUM("+ payment.formatAsString()+") * 11%) - SUM("+ leedsDiscount.formatAsString()+")");
        } else if (k == 14) {
            formulaList.add("TRUNC(((SUM("+realPayment.formatAsString()+") + " + deliveryText + ") * 2.65%),0) + ROUND(((TRUNC(((SUM("+ realPayment.formatAsString()+ ") + "+ deliveryText +") * 2.65%),0)) *10%), 0)");
        } else if (k == 15) {
            formulaList.add("(SUM("+ payment.formatAsString()+") * 11%) - SUM("+ leedsDiscount.formatAsString()+") + " + "TRUNC(((SUM("+realPayment.formatAsString()+") + " + deliveryText + ") * 2.65%),0) + ROUND(((TRUNC(((SUM("+ realPayment.formatAsString()+ ") + "+ deliveryText +") * 2.65%),0)) *10%), 0)");
        } else if (k == 16) {
            formulaList.add("SUM("+ realPayment.formatAsString()+") + " + deliveryText + " - " + "((SUM("+ payment.formatAsString()+") * 11%) - SUM("+ leedsDiscount.formatAsString()+") + " + "TRUNC(((SUM("+realPayment.formatAsString()+") + " + deliveryText + ") * 2.65%),0) + ROUND(((TRUNC(((SUM("+ realPayment.formatAsString()+ ") + "+ deliveryText +") * 2.65%),0)) *10%), 0))");
        }
    }

    private void setMergedCell(Sheet sheet, List<CellRangeAddress> mergedRegions, List<String> formulaList) {
        int index = 0;
        for (CellRangeAddress mergedRegion : mergedRegions) {
            //단일 셀이 아닐 경우
            if(mergedRegion.getFirstRow() != mergedRegion.getLastRow()){
                sheet.addMergedRegion(mergedRegion);
            }
            // 수식 적용
            if(mergedRegion.getFirstColumn() != 11){
                sheet.getRow(mergedRegion.getFirstRow()).getCell(mergedRegion.getFirstColumn()).setCellFormula(formulaList.get(index));
                index++;
            }

        }
    }
}
