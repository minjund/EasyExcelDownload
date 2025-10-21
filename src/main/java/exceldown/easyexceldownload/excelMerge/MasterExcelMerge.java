package exceldown.easyexceldownload.excelMerge;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

import static exceldown.easyexceldownload.excelMerge.ExcelColumnIndex.Master;

/**
 * Excel merge handler for Master Excel operations
 * Handles both seller-level and order-level merges with formula calculations
 */
public class MasterExcelMerge extends AbstractExcelMerge {

    private CellRangeAddress paymentRange;
    private CellRangeAddress pgFeeTargetRange;
    private CellRangeAddress deliveryRange;
    private CellRangeAddress discountRange;

    /**
     * Merges cells by seller within orders
     */
    public void sellerMergeValue(Sheet sheet, int totalNumberOfRows, int[] mergeCellNum,
                                int[] formulaCellNum) {
        mergeCellsByTwoColumns(sheet, totalNumberOfRows, mergeCellNum, formulaCellNum,
                              Master.ORDER_ID, Master.SELLER_NAME);
    }

    /**
     * Merges cells by order ID
     */
    public void orderMergeValue(Sheet sheet, int totalNumberOfRows, int[] mergeCellNum,
                               int[] formulaCellNum) {
        mergeCellsByColumn(sheet, totalNumberOfRows, mergeCellNum, formulaCellNum,
                          Master.ORDER_ID);
    }

    @Override
    protected void generateFormula(Sheet sheet, int[] formulaCellNum, int columnIndex,
                                   int startRowIndex, int endRowIndex, List<String> formulaList) {
        initializeRanges(formulaCellNum, startRowIndex, endRowIndex);

        if (isOrderLevelColumn(columnIndex)) {
            generateOrderFormula(sheet, formulaCellNum, columnIndex, startRowIndex, formulaList);
        } else if (isSellerLevelColumn(columnIndex)) {
            generateSellerFormula(sheet, columnIndex, startRowIndex, formulaList);
        }
    }

    /**
     * Initializes cell ranges for formula calculations
     */
    private void initializeRanges(int[] formulaCellNum, int startRowIndex, int endRowIndex) {
        paymentRange = createCellRange(startRowIndex, endRowIndex, formulaCellNum[0]);
        pgFeeTargetRange = createCellRange(startRowIndex, endRowIndex, formulaCellNum[1]);
        deliveryRange = createCellRange(startRowIndex, endRowIndex, formulaCellNum[2]);
        discountRange = createCellRange(startRowIndex, endRowIndex, formulaCellNum[3]);
    }

    /**
     * Generates formulas for seller-level columns
     */
    private void generateSellerFormula(Sheet sheet, int columnIndex, int startRowIndex,
                                      List<String> formulaList) {
        String deliveryCell = getFirstCellReference(deliveryRange);

        switch (columnIndex) {
            case Master.PAYMENT_SUM:
                formulaList.add(sumFormula(paymentRange));
                break;
            case Master.DELIVERY_AMOUNT:
                convertToNumeric(sheet, startRowIndex, columnIndex);
                break;
            case Master.COMMISSION:
                formulaList.add(String.format("(SUM(%s)*11/100) - SUM(%s)",
                    formatRangeAsString(pgFeeTargetRange),
                    formatRangeAsString(discountRange)));
                break;
            case Master.PG_FEE:
                formulaList.add(calculatePgFee(paymentRange, deliveryCell));
                break;
            case Master.SETTLEMENT:
                formulaList.add(calculateSellerSettlement(paymentRange, deliveryCell));
                break;
        }
    }

    /**
     * Generates formulas for order-level columns
     */
    private void generateOrderFormula(Sheet sheet, int[] formulaCellNum, int columnIndex,
                                     int startRowIndex, List<String> formulaList) {
        CellRangeAddress totalCalculateAmountRange = createCellRange(startRowIndex,
            startRowIndex, formulaCellNum[4]);
        CellRangeAddress couponAmountRange = createCellRange(startRowIndex,
            startRowIndex, formulaCellNum[5]);
        CellRangeAddress totalFeeRange = createCellRange(startRowIndex,
            startRowIndex, formulaCellNum[6]);

        String couponCell = getFirstCellReference(couponAmountRange);
        String deliverySum = sumFormula(deliveryRange);
        String totalFeeSum = sumFormula(totalFeeRange);

        switch (columnIndex) {
            case Master.TOTAL_CALCULATE_AMOUNT:
                formulaList.add(sumFormula(totalCalculateAmountRange));
                break;
            case Master.ORDER_PAYMENT_SUM:
                formulaList.add(sumFormula(paymentRange));
                break;
            case Master.LEEDS_COUPON:
                convertToNumeric(sheet, startRowIndex, columnIndex);
                break;
            case Master.ORDER_DELIVERY_AMOUNT:
                formulaList.add(deliverySum);
                break;
            case Master.ORDER_TOTAL_PAYMENT:
                formulaList.add(String.format("SUM(%s) - %s + %s",
                    formatRangeAsString(paymentRange), couponCell, deliverySum));
                break;
            case Master.ORDER_PG_FEE:
                formulaList.add(calculateOrderPgFee(paymentRange, couponCell, deliverySum));
                break;
            case Master.ORDER_SETTLEMENT:
                formulaList.add(calculateOrderSettlement(paymentRange, couponCell, deliverySum));
                break;
            case Master.ORDER_TOTAL_FEE:
                formulaList.add(String.format("%s - %s",
                    calculateOrderPgFee(paymentRange, couponCell, deliverySum), totalFeeSum));
                break;
            case Master.ORDER_FINAL_SETTLEMENT:
                formulaList.add(String.format("%s - SUM(%s)",
                    calculateOrderSettlement(paymentRange, couponCell, deliverySum),
                    formatRangeAsString(totalCalculateAmountRange)));
                break;
        }
    }

    /**
     * Calculates PG fee formula for seller level
     */
    private String calculatePgFee(CellRangeAddress paymentRange, String deliveryCell) {
        String baseAmount = String.format("SUM(%s) + %s",
            formatRangeAsString(paymentRange), deliveryCell);
        String pgFee = String.format("TRUNC((%s * 2.65%%),0)", baseAmount);
        String vatOnPgFee = String.format("ROUND((%s *10%%), 0)", pgFee);
        return String.format("%s + %s", pgFee, vatOnPgFee);
    }

    /**
     * Calculates settlement formula for seller level
     */
    private String calculateSellerSettlement(CellRangeAddress paymentRange, String deliveryCell) {
        String totalAmount = String.format("SUM(%s) + %s",
            formatRangeAsString(paymentRange), deliveryCell);
        String commission = String.format("(SUM(%s)*11/100) - SUM(%s)",
            formatRangeAsString(pgFeeTargetRange), formatRangeAsString(discountRange));
        String pgFee = calculatePgFee(paymentRange, deliveryCell);
        return String.format("%s - %s - %s", totalAmount, commission, pgFee);
    }

    /**
     * Calculates PG fee formula for order level
     */
    private String calculateOrderPgFee(CellRangeAddress paymentRange, String couponCell,
                                      String deliverySum) {
        String baseAmount = String.format("SUM(%s) - %s + %s",
            formatRangeAsString(paymentRange), couponCell, deliverySum);
        String pgFee = String.format("TRUNC((%s * 2.65%%),0)", baseAmount);
        String vatOnPgFee = String.format("ROUND((%s *10%%), 0)", pgFee);
        return String.format("%s + %s", pgFee, vatOnPgFee);
    }

    /**
     * Calculates settlement formula for order level
     */
    private String calculateOrderSettlement(CellRangeAddress paymentRange, String couponCell,
                                           String deliverySum) {
        String totalPayment = String.format("SUM(%s) - %s + %s",
            formatRangeAsString(paymentRange), couponCell, deliverySum);
        String pgFee = calculateOrderPgFee(paymentRange, couponCell, deliverySum);
        return String.format("%s - %s", totalPayment, pgFee);
    }

    /**
     * Creates a SUM formula for a cell range
     */
    private String sumFormula(CellRangeAddress range) {
        return String.format("SUM(%s)", formatRangeAsString(range));
    }

    /**
     * Converts a cell value from string to numeric type
     */
    private void convertToNumeric(Sheet sheet, int rowIndex, int columnIndex) {
        String value = sheet.getRow(rowIndex).getCell(columnIndex).getStringCellValue();
        sheet.getRow(rowIndex).getCell(columnIndex).setCellValue(Integer.parseInt(value));
        sheet.getRow(rowIndex).getCell(columnIndex).setCellType(CellType.NUMERIC);
    }

    /**
     * Checks if a column is order-level (requires special order-level formulas)
     */
    private boolean isOrderLevelColumn(int columnIndex) {
        return columnIndex >= Master.TOTAL_CALCULATE_AMOUNT;
    }

    /**
     * Checks if a column is seller-level
     */
    private boolean isSellerLevelColumn(int columnIndex) {
        return columnIndex >= Master.PAYMENT_SUM && columnIndex <= Master.SETTLEMENT;
    }

    @Override
    protected boolean isFormulaExcluded(int columnIndex) {
        return columnIndex == Master.DELIVERY_AMOUNT || columnIndex == Master.LEEDS_COUPON;
    }
}
