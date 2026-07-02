package com.easyexcel2poi.demo;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.ByteArrayInputStream;

import com.easyexcel2poi.excel.PoiExcel;
import com.easyexcel2poi.excel.WriteOptions;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockHttpServletResponse;

class PoiWebDemoTest {
    @Test
    void downloadWritesWorkbookToServletResponseStream() throws Exception {
        MockHttpServletResponse response = new MockHttpServletResponse();

        PoiExcel.write(response.getOutputStream(), PoiWriteDemoTest.DemoData.class,
            PoiWriteDemoTest.demoData(), WriteOptions.builder().sheetName("模板").build());

        assertThat(response.getContentAsByteArray()).isNotEmpty();
        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(response.getContentAsByteArray()))) {
            assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo("字符串标题");
            assertThat(workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue()).isEqualTo("字符串0");
        }
    }
}
