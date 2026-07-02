package com.easyexcel2poi.demo;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.net.URL;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

import com.easyexcel2poi.excel.ExcelColumn;
import com.easyexcel2poi.excel.ExcelColumnWidth;
import com.easyexcel2poi.excel.ExcelDatePattern;
import com.easyexcel2poi.excel.ExcelIgnoreColumn;
import com.easyexcel2poi.excel.ExcelLoopMerge;
import com.easyexcel2poi.excel.ExcelRowHeight;
import com.easyexcel2poi.excel.ImageCellData;
import com.easyexcel2poi.excel.PoiExcel;
import com.easyexcel2poi.excel.WriteCellData;
import com.easyexcel2poi.excel.WriteOptions;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class PoiWriteDemoTest {
    @Test
    void simpleWriteMatchesReferenceSnapshot() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        PoiExcel.write(outputStream, DemoData.class, demoData(), WriteOptions.builder().sheetName("模板").build());

        assertWorkbookMatches(outputStream, "easyexcel-baseline/write-simple.snapshot");
    }

    @Test
    void complexHeadAndLoopMergeMatchReferenceSnapshot() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        PoiExcel.write(outputStream, MergeData.class, mergeData(),
            WriteOptions.builder().sheetName("模板").build());

        assertWorkbookMatches(outputStream, "easyexcel-baseline/write-merge.snapshot");
    }

    @Test
    void imageAndCellDataWriteMatchesReferenceSnapshot() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        String imagePath = resourceFile("demo/img.jpg").getAbsolutePath();
        URL imageUrl = resourceFile("demo/img.jpg").toURI().toURL();

        ImageDemoData data = new ImageDemoData();
        data.setByteArray(readResourceBytes("demo/img.jpg"));
        data.setFile(resourceFile("demo/img.jpg"));
        data.setString(imagePath);
        data.setInputStream(new ByteArrayInputStream(readResourceBytes("demo/img.jpg")));
        data.setUrl(imageUrl);
        WriteCellData cellData = new WriteCellData("额外的放一些文字");
        cellData.addImage(new ImageCellData(readResourceBytes("demo/img.jpg"), 0, 0, 0, 0));
        cellData.addImage(new ImageCellData(readResourceBytes("demo/img.jpg"), 0, 0, 0, 1));
        data.setWriteCellDataFile(cellData);

        PoiExcel.write(outputStream, ImageDemoData.class, singleton(data),
            WriteOptions.builder().sheetName("模板").build());

        assertWorkbookMatches(outputStream, "easyexcel-baseline/write-image.snapshot");
    }

    private static void assertWorkbookMatches(ByteArrayOutputStream outputStream, String baseline) throws Exception {
        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(outputStream.toByteArray()))) {
            assertThat(WorkbookSnapshot.snapshot(workbook)).isEqualTo(WorkbookSnapshot.resourceText(baseline));
        }
    }

    static List<DemoData> demoData() {
        List<DemoData> list = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(LocalDateTime.of(2020, 1, i + 1, 1, 1, 1));
            data.setDoubleData((double)i + 1);
            data.setIgnore("ignore-" + i);
            list.add(data);
        }
        return list;
    }

    private static List<MergeData> mergeData() {
        List<MergeData> list = new ArrayList<>();
        for (int i = 0; i < 4; i++) {
            MergeData data = new MergeData();
            data.setString("合并" + (i / 2));
            data.setDate(LocalDateTime.of(2020, 2, i + 1, 1, 1, 1));
            data.setDoubleData((double)i + 1);
            list.add(data);
        }
        return list;
    }

    private static <T> List<T> singleton(T value) {
        List<T> values = new ArrayList<>();
        values.add(value);
        return values;
    }

    private static File resourceFile(String name) {
        URL url = PoiWriteDemoTest.class.getClassLoader().getResource(name);
        if (url == null) {
            throw new IllegalStateException("Missing resource " + name);
        }
        return new File(url.getFile());
    }

    private static byte[] readResourceBytes(String name) throws Exception {
        try (InputStream inputStream = PoiWriteDemoTest.class.getClassLoader().getResourceAsStream(name)) {
            if (inputStream == null) {
                throw new IllegalStateException("Missing resource " + name);
            }
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            byte[] buffer = new byte[1024];
            int len;
            while ((len = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, len);
            }
            return outputStream.toByteArray();
        }
    }

    public static class DemoData {
        @ExcelColumn("字符串标题")
        private String string;
        @ExcelColumn("日期标题")
        @ExcelDatePattern("yyyy-MM-dd HH:mm:ss")
        private LocalDateTime date;
        @ExcelColumn("数字标题")
        private Double doubleData;
        @ExcelIgnoreColumn
        private String ignore;

        public String getString() {
            return string;
        }

        public void setString(String string) {
            this.string = string;
        }

        public LocalDateTime getDate() {
            return date;
        }

        public void setDate(LocalDateTime date) {
            this.date = date;
        }

        public Double getDoubleData() {
            return doubleData;
        }

        public void setDoubleData(Double doubleData) {
            this.doubleData = doubleData;
        }

        public void setIgnore(String ignore) {
            this.ignore = ignore;
        }
    }

    public static class MergeData {
        @ExcelColumn({"主标题", "字符串标题"})
        @ExcelLoopMerge(eachRow = 2)
        private String string;
        @ExcelColumn({"主标题", "日期标题"})
        @ExcelDatePattern("yyyy-MM-dd HH:mm:ss")
        private LocalDateTime date;
        @ExcelColumn({"主标题", "数字标题"})
        private Double doubleData;

        public void setString(String string) {
            this.string = string;
        }

        public void setDate(LocalDateTime date) {
            this.date = date;
        }

        public void setDoubleData(Double doubleData) {
            this.doubleData = doubleData;
        }
    }

    @ExcelRowHeight(head = 20, content = 80)
    @ExcelColumnWidth(16)
    public static class ImageDemoData {
        @ExcelColumn("byteArray")
        private byte[] byteArray;
        @ExcelColumn("file")
        private File file;
        @ExcelColumn("string")
        private String string;
        @ExcelColumn("inputStream")
        private InputStream inputStream;
        @ExcelColumn("url")
        private URL url;
        @ExcelColumn("writeCellDataFile")
        private WriteCellData writeCellDataFile;

        public void setByteArray(byte[] byteArray) {
            this.byteArray = byteArray;
        }

        public void setFile(File file) {
            this.file = file;
        }

        public void setString(String string) {
            this.string = string;
        }

        public void setInputStream(InputStream inputStream) {
            this.inputStream = inputStream;
        }

        public void setUrl(URL url) {
            this.url = url;
        }

        public void setWriteCellDataFile(WriteCellData writeCellDataFile) {
            this.writeCellDataFile = writeCellDataFile;
        }
    }
}
