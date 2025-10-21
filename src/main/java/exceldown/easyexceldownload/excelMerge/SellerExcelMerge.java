package exceldown.easyexceldownload.excelMerge;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

import static exceldown.easyexceldownload.excelMerge.ExcelColumnIndex.Seller;

/**
 * Excel merge handler for Seller Excel operations
 * Handles seller-level merges with formula calculations
 */
public class SellerExcelMerge extends AbstractExcelMerge {

    private CellRangeAddress paymentRange;
    private CellRangeAddress leedsDiscountRange;
    private CellRangeAddress realPaymentRange;
    private CellRangeAddress deliveryRange;

    /**
     * Merges cells by order ID
     */
    public void sellerMergeValue(Sheet sheet, int totalNumberOfRows, int[] mergeCellNum,
                                int[] formulaCellNum) {
        mergeCellsByColumn(sheet, totalNumberOfRows, mergeCellNum, formulaCellNum,
                          Seller.ORDER_ID);
    }

    @Override
    protected void generateFormula(Sheet sheet, int[] formulaCellNum, int columnIndex,
                                   int startRowIndex, int endRowIndex, List<String> formulaList) {
        initializeRanges(formulaCellNum, startRowIndex, endRowIndex);

        String deliveryCell = getFirstCellReference(deliveryRange);

        switch (columnIndex) {
            case Seller.REAL_PAYMENT_SUM:
                formulaList.add(sumFormula(realPaymentRange));
                break;
            case Seller.LEEDS_COUPON:
                convertToNumeric(sheet, startRowIndex, columnIndex);
                break;
            case Seller.COMMISSION:
                formulaList.add(calculateCommission());
                break;
            case Seller.PG_FEE:
                formulaList.add(calculatePgFee(deliveryCell));
                break;
            case Seller.TOTAL_FEE:
                formulaList.add(calculateTotalFee(deliveryCell));
                break;
            case Seller.SETTLEMENT:
                formulaList.add(calculateSettlement(deliveryCell));
                break;
        }
    }

    /**
     * Initializes cell ranges for formula calculations
     */
    private void initializeRanges(int[] formulaCellNum, int startRowIndex, int endRowIndex) {
        paymentRange = createCellRange(startRowIndex, endRowIndex, formulaCellNum[0]);
        leedsDiscountRange = createCellRange(startRowIndex, endRowIndex, formulaCellNum[1]);
        realPaymentRange = createCellRange(startRowIndex, endRowIndex, formulaCellNum[2]);
        deliveryRange = createCellRange(startRowIndex, endRowIndex, formulaCellNum[3]);
    }

    /**
     * Calculates commission formula
     */
    private String calculateCommission() {
        return String.format("(SUM(%s) * 11%%) - SUM(%s)",
            formatRangeAsString(paymentRange),
            formatRangeAsString(leedsDiscountRange));
    }

    /**
     * Calculates PG fee formula
     */
    private String calculatePgFee(String deliveryCell) {
        String baseAmount = String.format("SUM(%s) + %s",
            formatRangeAsString(realPaymentRange), deliveryCell);
        String pgFee = String.format("TRUNC((%s * 2.65%%),0)", baseAmount);
        String vatOnPgFee = String.format("ROUND((%s *10%%), 0)", pgFee);
        return String.format("%s + %s", pgFee, vatOnPgFee);
    }

    /**
     * Calculates total fee formula (commission + PG fee)
     */
    private String calculateTotalFee(String deliveryCell) {
        String commission = calculateCommission();
        String pgFee = calculatePgFee(deliveryCell);
        return String.format("%s + %s", commission, pgFee);
    }

    /**
     * Calculates settlement formula
     */
    private String calculateSettlement(String deliveryCell) {
        String totalAmount = String.format("SUM(%s) + %s",
            formatRangeAsString(realPaymentRange), deliveryCell);
        String totalFee = calculateTotalFee(deliveryCell);
        return String.format("%s - (%s)", totalAmount, totalFee);
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

    @Override
    protected boolean isFormulaExcluded(int columnIndex) {
        return columnIndex == Seller.REAL_PAYMENT_SUM;
    }
}
