package com.easyexcel2poi.excel;

import org.apache.poi.ss.usermodel.Sheet;

public interface SheetWriteHandler {
    void afterSheetCreate(Sheet sheet);
}
