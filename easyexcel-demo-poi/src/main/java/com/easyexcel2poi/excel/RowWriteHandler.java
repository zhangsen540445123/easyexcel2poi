package com.easyexcel2poi.excel;

import org.apache.poi.ss.usermodel.Row;

public interface RowWriteHandler {
    void afterRowDispose(Row row, int rowIndex, boolean head);
}
