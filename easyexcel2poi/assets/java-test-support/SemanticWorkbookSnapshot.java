package com.easyexcel2poi.testsupport;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public final class SemanticWorkbookSnapshot {
    private SemanticWorkbookSnapshot() {
    }

    public static String snapshot(Workbook workbook) {
        DataFormatter formatter = new DataFormatter();
        StringBuilder out = new StringBuilder();
        out.append("sheets=").append(workbook.getNumberOfSheets()).append('\n');
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            out.append("sheet[").append(sheetIndex).append("]=").append(sheet.getSheetName()).append('\n');
            out.append("merged=").append(mergedRegions(sheet)).append('\n');
            out.append("pictures=").append(pictureCount(sheet)).append('\n');
            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }
                List<String> cells = new ArrayList<String>();
                int lastCell = Math.max(row.getLastCellNum(), (short)0);
                for (int cellIndex = 0; cellIndex < lastCell; cellIndex++) {
                    cells.add(cellSnapshot(row.getCell(cellIndex), formatter));
                }
                out.append("row[").append(rowIndex).append("]=").append(String.join("|", cells)).append('\n');
            }
        }
        return out.toString().replace("\r\n", "\n");
    }

    private static String cellSnapshot(Cell cell, DataFormatter formatter) {
        if (cell == null) {
            return "BLANK:";
        }
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            return "FORMULA:" + cell.getCellFormula();
        }
        String comment = "";
        Comment cellComment = cell.getCellComment();
        if (cellComment != null) {
            comment = ";comment=" + cellComment.getString().getString();
        }
        String hyperlink = "";
        Hyperlink cellHyperlink = cell.getHyperlink();
        if (cellHyperlink != null) {
            hyperlink = ";hyperlink=" + cellHyperlink.getAddress();
        }
        return type.name() + ":" + formatter.formatCellValue(cell) + comment + hyperlink;
    }

    private static String mergedRegions(Sheet sheet) {
        List<String> ranges = new ArrayList<String>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            ranges.add(region.formatAsString());
        }
        return ranges.toString();
    }

    private static int pictureCount(Sheet sheet) {
        if (sheet instanceof XSSFSheet) {
            Drawing<?> drawing = ((XSSFSheet)sheet).getDrawingPatriarch();
            if (drawing instanceof XSSFDrawing) {
                return ((XSSFDrawing)drawing).getShapes().size();
            }
        }
        return 0;
    }
}
