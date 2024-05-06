package exceldown.easyexceldownload.excelMerge;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class MasterExcelMerge {

    private CellRangeAddress payment = new CellRangeAddress(0,0,0,0);
    private CellRangeAddress pgFeeTargetMerge = new CellRangeAddress(0,0,0,0);
    private CellRangeAddress deliveryMerge = new CellRangeAddress(0,0,0,0);
    private CellRangeAddress discount = new CellRangeAddress(0,0,0,0);

    private void defaultSetting(int[] formulaCellNum, int startRowIndex, int endRowIndex) {
        payment = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[0], formulaCellNum[0]);
        pgFeeTargetMerge = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[1], formulaCellNum[1]);
        deliveryMerge = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[2], formulaCellNum[2]);
        discount = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[3], formulaCellNum[3]);
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
            String orderIdValue = currentRow.getCell(1).getStringCellValue();
            String prevOrderIdValue = i > 0 ? sheet.getRow(i - 1).getCell(1).getStringCellValue() : null;

            // 셀러 현재 ROW value
            String currentAValue = currentRow.getCell(2).getStringCellValue();

            // 셀러 이전 ROW value
            String prevAValue = i > 0 ? sheet.getRow(i - 1).getCell(2).getStringCellValue() : null;

            if(orderIdValue.equals(prevOrderIdValue) && currentAValue.equals(prevAValue)) {
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


    public void orderMergeValue(Sheet sheet, int totalNumberOfRows, int[] mergeCellNum, int[] formulaCellNum) {
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        List<String> formulaList = new ArrayList<>();

        int startRowIndex = 0;
        int endRowIndex = 0;

        for (int i = 0; i <= totalNumberOfRows; i++) {
            Row currentRow = sheet.getRow(i);
            if (currentRow == null) continue; // 현재 행이 없으면 건너뜁니다.

            //주문번호 value
            String orderIdValue = currentRow.getCell(1).getStringCellValue();
            String prevOrderIdValue = i > 0 ? sheet.getRow(i - 1).getCell(1).getStringCellValue() : null;

            if(orderIdValue.equals(prevOrderIdValue)) {
                endRowIndex = i;
            } else {
                if(startRowIndex < endRowIndex) {
                    for (int k : mergeCellNum) {
                        CellRangeAddress range = new CellRangeAddress(startRowIndex, endRowIndex, k, k);
                        setOrderFormula(sheet, formulaCellNum, k, startRowIndex, endRowIndex, formulaList);
                        mergedRegions.add(range);
                    }
                }

                if(startRowIndex > endRowIndex) {
                    for (int k : mergeCellNum) {
                        CellRangeAddress range = new CellRangeAddress(startRowIndex, startRowIndex, k, k);
                        setOrderFormula(sheet, formulaCellNum, k, startRowIndex, startRowIndex, formulaList);
                        mergedRegions.add(range);
                    }
                }

                startRowIndex = i;
            }
        }

        for (int i : mergeCellNum) {

            if (startRowIndex < endRowIndex) {
                CellRangeAddress range = new CellRangeAddress(startRowIndex, endRowIndex, i, i);
                setOrderFormula(sheet, formulaCellNum, i, startRowIndex, endRowIndex, formulaList);
                mergedRegions.add(range);
            } else {
                CellRangeAddress range = new CellRangeAddress(startRowIndex, startRowIndex, i, i);
                setOrderFormula(sheet, formulaCellNum, i, startRowIndex, startRowIndex, formulaList);
                mergedRegions.add(range);
            }


        }

        setMergedCell(sheet, mergedRegions, formulaList);
    }

    private void setOrderFormula(Sheet sheet, int[] formulaCellNum, int k, int startRowIndex, int endRowIndex, List<String> formulaList) {
        defaultSetting(formulaCellNum, startRowIndex, endRowIndex);

        CellRangeAddress totalCalculateAmount = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[4], formulaCellNum[4]);
        CellRangeAddress couponAmount = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[5], formulaCellNum[5]);
        CellRangeAddress totalFee = new CellRangeAddress(startRowIndex, endRowIndex, formulaCellNum[6], formulaCellNum[6]);

        String couponText = couponAmount.formatAsString().split(":")[0];
        String deliveryAmount = "(SUM("+deliveryMerge.formatAsString()+"))";
        String totalFeeAmount = "(SUM("+totalFee.formatAsString()+"))";

        generateFormula(sheet, k, startRowIndex, formulaList, totalCalculateAmount, deliveryAmount, couponText, totalFeeAmount);
    }

    private void generateFormula(Sheet sheet, int k, int startRowIndex, List<String> formulaList, CellRangeAddress totalCalculateAmount, String deliveryAmount, String couponText, String totalFeeAmount) {
        if(k == 17){
            formulaList.add("SUM("+ totalCalculateAmount.formatAsString()+")");
        } else if (k == 18) {
            formulaList.add("SUM("+payment.formatAsString()+")");
        } else if (k == 19) {
            String leedsCouponAmount  = sheet.getRow(startRowIndex).getCell(19).getStringCellValue();
            sheet.getRow(startRowIndex).getCell(19).setCellValue(Integer.parseInt(leedsCouponAmount));
            sheet.getRow(startRowIndex).getCell(19).setCellType(CellType.NUMERIC);
        } else if (k == 20) {
            formulaList.add(deliveryAmount);
        } else if (k == 21) {
            formulaList.add("SUM("+payment.formatAsString()+") - "+ couponText + " + " + deliveryAmount);
        } else if (k == 22) {
            formulaList.add("TRUNC(((SUM("+payment.formatAsString()+") - "+ couponText +" + "+ deliveryAmount +") * 2.65%),0) + ROUND(((TRUNC(((SUM("+ payment.formatAsString()+ ") - "+ couponText +" + "+ deliveryAmount +") * 2.65%),0)) *10%), 0)");
        } else if (k == 23) {
            formulaList.add("(SUM("+payment.formatAsString()+") - "+ couponText +" + "+ deliveryAmount +") - (TRUNC(((SUM("+payment.formatAsString()+") - "+ couponText +" + "+ deliveryAmount + ") * 2.65%),0) + ROUND(((TRUNC(((SUM("+payment.formatAsString()+") - "+ couponText +" + " + deliveryAmount + ") * 2.65%),0)) *10%), 0))");
        } else if (k == 24) {
            formulaList.add("TRUNC(((SUM("+payment.formatAsString()+") - "+ couponText +" + "+ deliveryAmount +") * 2.65%),0) + ROUND(((TRUNC(((SUM("+ payment.formatAsString()+ ") - "+ couponText +" + "+ deliveryAmount +") * 2.65%),0)) *10%), 0) - " + totalFeeAmount);
        } else if (k == 25) {
            formulaList.add("(SUM("+payment.formatAsString()+") - "+ couponText +" + "+ deliveryAmount +") - (TRUNC(((SUM("+payment.formatAsString()+") - "+ couponText +" + "+ deliveryAmount + ") * 2.65%),0) + ROUND(((TRUNC(((SUM("+payment.formatAsString()+") - "+ couponText +" + " + deliveryAmount + ") * 2.65%),0)) *10%), 0)) - SUM("+ totalCalculateAmount.formatAsString()+")");
        }
    }

    private void generateFormula(Sheet sheet, int k, int startRowIndex, List<String> formulaList, String deliveryText) {
        if(k == 12){
            formulaList.add("SUM("+payment.formatAsString()+")");
        } else if (k == 13) {
            String deliveryAmount  = sheet.getRow(startRowIndex).getCell(13).getStringCellValue();
            sheet.getRow(startRowIndex).getCell(13).setCellValue(Integer.parseInt(deliveryAmount));
            sheet.getRow(startRowIndex).getCell(13).setCellType(CellType.NUMERIC);
        } else if (k == 14) {
            formulaList.add("(SUM("+pgFeeTargetMerge.formatAsString()+")*11/100) - SUM("+discount.formatAsString()+")");
        } else if (k == 15) {
            formulaList.add("TRUNC(((SUM("+payment.formatAsString()+") + "+ deliveryText +") * 2.65%),0) + ROUND(((TRUNC(((SUM("+ payment.formatAsString()+ ") + " + deliveryText + ") * 2.65%),0)) *10%), 0)");
        } else if (k == 16) {
            formulaList.add("(SUM("+payment.formatAsString()+")" + " + " + deliveryText + ") - ((SUM("+pgFeeTargetMerge.formatAsString()+")*11/100) - SUM("+discount.formatAsString()+")) - (TRUNC(((SUM("+payment.formatAsString()+") + "+ deliveryText +") * 2.65%),0) + ROUND(((TRUNC(((SUM("+ payment.formatAsString()+ ") + " + deliveryText + ") * 2.65%),0)) *10%), 0))");
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
            if(mergedRegion.getFirstColumn() != 13 && mergedRegion.getFirstColumn() != 19){
                sheet.getRow(mergedRegion.getFirstRow()).getCell(mergedRegion.getFirstColumn()).setCellFormula(formulaList.get(index));
                index++;
            }

        }
    }
}
