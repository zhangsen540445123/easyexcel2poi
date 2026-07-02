package com.easyexcel2poi.excel;

import org.apache.poi.ss.usermodel.Cell;

public interface CellWriteHandler {
    void afterCellDispose(Cell cell, int rowIndex, int columnIndex, boolean head);
}
