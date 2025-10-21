package exceldown.easyexceldownload.excelMerge;

/**
 * Excel column index constants
 * This class defines all column indices used in Excel merge operations
 */
public final class ExcelColumnIndex {

    private ExcelColumnIndex() {
        // Prevent instantiation
    }

    // Seller Excel Columns
    public static final class Seller {
        public static final int ORDER_ID = 0;
        public static final int REAL_PAYMENT_SUM = 11;
        public static final int LEEDS_COUPON = 12;
        public static final int COMMISSION = 13;
        public static final int PG_FEE = 14;
        public static final int TOTAL_FEE = 15;
        public static final int SETTLEMENT = 16;
    }

    // Master Excel Columns
    public static final class Master {
        public static final int ORDER_ID = 1;
        public static final int SELLER_NAME = 2;
        public static final int PAYMENT_SUM = 12;
        public static final int DELIVERY_AMOUNT = 13;
        public static final int COMMISSION = 14;
        public static final int PG_FEE = 15;
        public static final int SETTLEMENT = 16;
        public static final int TOTAL_CALCULATE_AMOUNT = 17;
        public static final int ORDER_PAYMENT_SUM = 18;
        public static final int LEEDS_COUPON = 19;
        public static final int ORDER_DELIVERY_AMOUNT = 20;
        public static final int ORDER_TOTAL_PAYMENT = 21;
        public static final int ORDER_PG_FEE = 22;
        public static final int ORDER_SETTLEMENT = 23;
        public static final int ORDER_TOTAL_FEE = 24;
        public static final int ORDER_FINAL_SETTLEMENT = 25;
    }
}
