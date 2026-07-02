package com.easyexcel2poi.excel;

import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;

public final class WriteOptions {
    private final String sheetName;
    private final boolean needHead;
    private final Set<String> includeFields;
    private final Set<String> excludeFields;

    private WriteOptions(Builder builder) {
        this.sheetName = builder.sheetName;
        this.needHead = builder.needHead;
        this.includeFields = Collections.unmodifiableSet(new LinkedHashSet<String>(builder.includeFields));
        this.excludeFields = Collections.unmodifiableSet(new LinkedHashSet<String>(builder.excludeFields));
    }

    public static Builder builder() {
        return new Builder();
    }

    public String getSheetName() {
        return sheetName;
    }

    public boolean isNeedHead() {
        return needHead;
    }

    public Set<String> getIncludeFields() {
        return includeFields;
    }

    public Set<String> getExcludeFields() {
        return excludeFields;
    }

    public static final class Builder {
        private String sheetName = "Sheet1";
        private boolean needHead = true;
        private final Set<String> includeFields = new LinkedHashSet<String>();
        private final Set<String> excludeFields = new LinkedHashSet<String>();

        public Builder sheetName(String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

        public Builder needHead(boolean needHead) {
            this.needHead = needHead;
            return this;
        }

        public Builder includeField(String fieldName) {
            this.includeFields.add(fieldName);
            return this;
        }

        public Builder excludeField(String fieldName) {
            this.excludeFields.add(fieldName);
            return this;
        }

        public WriteOptions build() {
            return new WriteOptions(this);
        }
    }
}
