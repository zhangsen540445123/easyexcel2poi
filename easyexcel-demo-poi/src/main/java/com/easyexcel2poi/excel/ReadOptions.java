package com.easyexcel2poi.excel;

public final class ReadOptions {
    private final int sheetIndex;
    private final int headRowNumber;

    private ReadOptions(Builder builder) {
        this.sheetIndex = builder.sheetIndex;
        this.headRowNumber = builder.headRowNumber;
    }

    public static Builder builder() {
        return new Builder();
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public int getHeadRowNumber() {
        return headRowNumber;
    }

    public static final class Builder {
        private int sheetIndex = 0;
        private int headRowNumber = 1;

        public Builder sheetIndex(int sheetIndex) {
            this.sheetIndex = sheetIndex;
            return this;
        }

        public Builder headRowNumber(int headRowNumber) {
            this.headRowNumber = headRowNumber;
            return this;
        }

        public ReadOptions build() {
            return new ReadOptions(this);
        }
    }
}
