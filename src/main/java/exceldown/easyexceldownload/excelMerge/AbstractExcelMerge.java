package exceldown.easyexceldownload.excelMerge;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * Abstract base class for Excel merge operations
 * Provides common functionality for merging cells and applying formulas
 */
public abstract class AbstractExcelMerge {

    /**
     * Merges cells based on matching values in the specified column
     *
     * @param sheet The Excel sheet to process
     * @param totalNumberOfRows Total number of rows to process
     * @param mergeCellNum Array of column indices to merge
     * @param formulaCellNum Array of column indices for formula cells
     * @param groupingColumnIndex Column index to use for grouping rows
     */
    protected void mergeCellsByColumn(Sheet sheet, int totalNumberOfRows, int[] mergeCellNum,
                                      int[] formulaCellNum, int groupingColumnIndex) {
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        List<String> formulaList = new ArrayList<>();
        int startRowIndex = 0;
        int endRowIndex = 0;

        for (int i = 0; i <= totalNumberOfRows; i++) {
            Row currentRow = sheet.getRow(i);
            if (currentRow == null) {
                continue;
            }

            String currentValue = currentRow.getCell(groupingColumnIndex).getStringCellValue();
            String previousValue = i > 0 ? sheet.getRow(i - 1).getCell(groupingColumnIndex).getStringCellValue() : null;

            if (currentValue.equals(previousValue)) {
                endRowIndex = i;
            } else {
                processMergeRange(sheet, mergeCellNum, formulaCellNum, mergedRegions,
                                formulaList, startRowIndex, endRowIndex);
                startRowIndex = i;
            }
        }

        // Process final range
        processMergeRange(sheet, mergeCellNum, formulaCellNum, mergedRegions,
                        formulaList, startRowIndex, endRowIndex);

        applyMergedCells(sheet, mergedRegions, formulaList);
    }

    /**
     * Merges cells based on matching values in two columns (for nested grouping)
     *
     * @param sheet The Excel sheet to process
     * @param totalNumberOfRows Total number of rows to process
     * @param mergeCellNum Array of column indices to merge
     * @param formulaCellNum Array of column indices for formula cells
     * @param primaryColumnIndex Primary column index for grouping
     * @param secondaryColumnIndex Secondary column index for grouping
     */
    protected void mergeCellsByTwoColumns(Sheet sheet, int totalNumberOfRows, int[] mergeCellNum,
                                         int[] formulaCellNum, int primaryColumnIndex,
                                         int secondaryColumnIndex) {
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        List<String> formulaList = new ArrayList<>();
        int startRowIndex = 0;
        int endRowIndex = 0;

        for (int i = 0; i <= totalNumberOfRows; i++) {
            Row currentRow = sheet.getRow(i);
            if (currentRow == null) {
                continue;
            }

            String currentPrimaryValue = currentRow.getCell(primaryColumnIndex).getStringCellValue();
            String previousPrimaryValue = i > 0 ? sheet.getRow(i - 1).getCell(primaryColumnIndex).getStringCellValue() : null;

            String currentSecondaryValue = currentRow.getCell(secondaryColumnIndex).getStringCellValue();
            String previousSecondaryValue = i > 0 ? sheet.getRow(i - 1).getCell(secondaryColumnIndex).getStringCellValue() : null;

            if (currentPrimaryValue.equals(previousPrimaryValue) &&
                currentSecondaryValue.equals(previousSecondaryValue)) {
                endRowIndex = i;
            } else {
                processMergeRange(sheet, mergeCellNum, formulaCellNum, mergedRegions,
                                formulaList, startRowIndex, endRowIndex);
                startRowIndex = i;
            }
        }

        // Process final range
        processMergeRange(sheet, mergeCellNum, formulaCellNum, mergedRegions,
                        formulaList, startRowIndex, endRowIndex);

        applyMergedCells(sheet, mergedRegions, formulaList);
    }

    /**
     * Processes a single merge range
     */
    private void processMergeRange(Sheet sheet, int[] mergeCellNum, int[] formulaCellNum,
                                   List<CellRangeAddress> mergedRegions, List<String> formulaList,
                                   int startRowIndex, int endRowIndex) {
        if (startRowIndex < endRowIndex) {
            addMergeRanges(sheet, mergeCellNum, formulaCellNum, mergedRegions,
                          formulaList, startRowIndex, endRowIndex);
        } else {
            addMergeRanges(sheet, mergeCellNum, formulaCellNum, mergedRegions,
                          formulaList, startRowIndex, startRowIndex);
        }
    }

    /**
     * Adds merge ranges for specified columns
     */
    private void addMergeRanges(Sheet sheet, int[] mergeCellNum, int[] formulaCellNum,
                               List<CellRangeAddress> mergedRegions, List<String> formulaList,
                               int startRowIndex, int endRowIndex) {
        for (int columnIndex : mergeCellNum) {
            CellRangeAddress range = new CellRangeAddress(startRowIndex, endRowIndex,
                                                          columnIndex, columnIndex);
            generateFormula(sheet, formulaCellNum, columnIndex, startRowIndex,
                          endRowIndex, formulaList);
            mergedRegions.add(range);
        }
    }

    /**
     * Applies merged cell regions and formulas to the sheet
     */
    private void applyMergedCells(Sheet sheet, List<CellRangeAddress> mergedRegions,
                                 List<String> formulaList) {
        int formulaIndex = 0;
        for (CellRangeAddress mergedRegion : mergedRegions) {
            // Only merge if not a single cell
            if (mergedRegion.getFirstRow() != mergedRegion.getLastRow()) {
                sheet.addMergedRegion(mergedRegion);
            }

            // Apply formula if not excluded
            if (!isFormulaExcluded(mergedRegion.getFirstColumn())) {
                sheet.getRow(mergedRegion.getFirstRow())
                     .getCell(mergedRegion.getFirstColumn())
                     .setCellFormula(formulaList.get(formulaIndex));
                formulaIndex++;
            }
        }
    }

    /**
     * Creates a cell range address for formula calculation
     */
    protected CellRangeAddress createCellRange(int startRow, int endRow, int column) {
        return new CellRangeAddress(startRow, endRow, column, column);
    }

    /**
     * Formats a cell range as a string for use in formulas
     */
    protected String formatRangeAsString(CellRangeAddress range) {
        return range.formatAsString();
    }

    /**
     * Gets the first cell reference from a range (e.g., "A1" from "A1:A10")
     */
    protected String getFirstCellReference(CellRangeAddress range) {
        return range.formatAsString().split(":")[0];
    }

    /**
     * Generates formula for a specific cell based on column index
     * Must be implemented by subclasses to provide specific formula logic
     */
    protected abstract void generateFormula(Sheet sheet, int[] formulaCellNum, int columnIndex,
                                           int startRowIndex, int endRowIndex,
                                           List<String> formulaList);

    /**
     * Determines if a column should be excluded from formula application
     * Can be overridden by subclasses
     */
    protected abstract boolean isFormulaExcluded(int columnIndex);
}
