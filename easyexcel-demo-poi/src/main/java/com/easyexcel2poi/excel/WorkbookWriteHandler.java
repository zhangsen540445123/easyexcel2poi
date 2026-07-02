package com.easyexcel2poi.excel;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookWriteHandler {
    void afterWorkbookCreate(Workbook workbook);

    void afterWorkbookDispose(Workbook workbook);
}
