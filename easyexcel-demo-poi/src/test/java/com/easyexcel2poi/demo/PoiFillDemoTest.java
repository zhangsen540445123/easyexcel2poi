package com.easyexcel2poi.demo;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import com.easyexcel2poi.excel.ExcelColumn;
import com.easyexcel2poi.excel.ExcelDatePattern;
import com.easyexcel2poi.excel.FillOptions;
import com.easyexcel2poi.excel.PoiExcel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class PoiFillDemoTest {
    @Test
    void fillTemplateMatchesReferenceSnapshot() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        Map<String, Object> scalars = new LinkedHashMap<>();
        scalars.put("total", 2);
        scalars.put("date", LocalDateTime.of(2020, 1, 1, 1, 1, 1));

        PoiExcel.fill(template(), outputStream, fillRows(), FillOptions.builder().scalars(scalars).build());

        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(outputStream.toByteArray()))) {
            assertThat(WorkbookSnapshot.snapshot(workbook))
                .isEqualTo(WorkbookSnapshot.resourceText("easyexcel-baseline/fill-simple.snapshot"));
        }
    }

    private static ByteArrayInputStream template() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("模板");
            sheet.createRow(0).createCell(0).setCellValue("日期");
            sheet.getRow(0).createCell(1).setCellValue("{date}");
            sheet.createRow(1).createCell(0).setCellValue("总数");
            sheet.getRow(1).createCell(1).setCellValue("{total}");
            sheet.createRow(2).createCell(0).setCellValue("名称");
            sheet.getRow(2).createCell(1).setCellValue("数字");
            sheet.getRow(2).createCell(2).setCellValue("时间");
            sheet.createRow(3).createCell(0).setCellValue("{.name}");
            sheet.getRow(3).createCell(1).setCellValue("{.number}");
            sheet.getRow(3).createCell(2).setCellValue("{.date}");
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return new ByteArrayInputStream(outputStream.toByteArray());
        }
    }

    private static List<FillData> fillRows() {
        List<FillData> list = new ArrayList<>();
        list.add(new FillData("张三", 5.2D, LocalDateTime.of(2020, 1, 1, 1, 1, 1)));
        list.add(new FillData("李四", 6.2D, LocalDateTime.of(2020, 1, 2, 1, 1, 1)));
        return list;
    }

    public static class FillData {
        @ExcelColumn("name")
        private final String name;
        @ExcelColumn("number")
        private final double number;
        @ExcelColumn("date")
        @ExcelDatePattern("yyyy-MM-dd HH:mm:ss")
        private final LocalDateTime date;

        FillData(String name, double number, LocalDateTime date) {
            this.name = name;
            this.number = number;
            this.date = date;
        }
    }
}
