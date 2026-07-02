package com.easyexcel2poi.excel;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;

public final class FillOptions {
    private final Map<String, Object> scalars;

    private FillOptions(Builder builder) {
        this.scalars = Collections.unmodifiableMap(new LinkedHashMap<String, Object>(builder.scalars));
    }

    public static Builder builder() {
        return new Builder();
    }

    public Map<String, Object> getScalars() {
        return scalars;
    }

    public static final class Builder {
        private final Map<String, Object> scalars = new LinkedHashMap<String, Object>();

        public Builder scalar(String name, Object value) {
            this.scalars.put(name, value);
            return this;
        }

        public Builder scalars(Map<String, Object> scalars) {
            this.scalars.putAll(scalars);
            return this;
        }

        public FillOptions build() {
            return new FillOptions(this);
        }
    }
}
