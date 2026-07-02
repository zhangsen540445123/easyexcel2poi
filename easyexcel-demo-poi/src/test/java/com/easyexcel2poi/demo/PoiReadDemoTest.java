package com.easyexcel2poi.demo;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.InputStream;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;

import com.easyexcel2poi.excel.ExcelColumn;
import com.easyexcel2poi.excel.PoiExcel;
import com.easyexcel2poi.excel.ReadOptions;

import org.junit.jupiter.api.Test;

class PoiReadDemoTest {
    @Test
    void simpleReadReturnsDemoRowsFromReferenceWorkbook() throws Exception {
        try (InputStream inputStream = resource("demo/demo.xlsx")) {
            List<ReadDemoData> rows = PoiExcel.read(inputStream, ReadDemoData.class, ReadOptions.builder().build());

            assertThat(rows).hasSize(10);
            assertThat(rows.get(0).getString()).isEqualTo("字符串0");
            assertThat(rows.get(0).getDate()).isEqualTo(LocalDateTime.of(2020, 1, 1, 1, 1, 1));
            assertThat(rows.get(0).getDoubleData()).isEqualTo(1.0D);
        }
    }

    @Test
    void noModelReadReturnsHeaderIndexedMaps() throws Exception {
        try (InputStream inputStream = resource("demo/demo.xlsx")) {
            List<Map<Integer, String>> rows = PoiExcel.readMaps(inputStream, ReadOptions.builder().headRowNumber(1).build());

            assertThat(rows).hasSize(10);
            assertThat(rows.get(0)).containsEntry(0, "字符串0");
            assertThat(rows.get(0)).containsEntry(2, "1");
        }
    }

    private static InputStream resource(String name) {
        InputStream inputStream = PoiReadDemoTest.class.getClassLoader().getResourceAsStream(name);
        if (inputStream == null) {
            throw new IllegalStateException("Missing resource " + name);
        }
        return inputStream;
    }

    public static class ReadDemoData {
        @ExcelColumn("字符串标题")
        private String string;
        @ExcelColumn("日期标题")
        private LocalDateTime date;
        @ExcelColumn("数字标题")
        private Double doubleData;

        public String getString() {
            return string;
        }

        public LocalDateTime getDate() {
            return date;
        }

        public Double getDoubleData() {
            return doubleData;
        }
    }
}
