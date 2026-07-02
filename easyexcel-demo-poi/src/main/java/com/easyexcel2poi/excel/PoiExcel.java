package com.easyexcel2poi.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.net.URL;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class PoiExcel {
    private PoiExcel() {
    }

    public static <T> void write(OutputStream outputStream, Class<T> headClass, List<T> data,
        WriteOptions options) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet(options.getSheetName());
            List<ColumnMeta> columns = columnsForWrite(headClass, options);
            int dataStartRow = 0;
            if (options.isNeedHead()) {
                dataStartRow = writeHead(workbook, sheet, headClass, columns);
            }
            writeRows(workbook, sheet, columns, data, dataStartRow);
            workbook.write(outputStream);
        } finally {
            workbook.close();
        }
    }

    public static <T> List<T> read(InputStream inputStream, Class<T> headClass, ReadOptions options)
        throws IOException {
        Workbook workbook = new XSSFWorkbook(inputStream);
        try {
            Sheet sheet = workbook.getSheetAt(options.getSheetIndex());
            List<ColumnMeta> columns = columnsForRead(headClass);
            Map<String, Integer> headerIndex = headerIndex(sheet, options.getHeadRowNumber());
            List<T> result = new ArrayList<T>();
            for (int rowIndex = options.getHeadRowNumber(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null || rowIsEmpty(row)) {
                    continue;
                }
                result.add(readBean(row, columns, headerIndex, headClass));
            }
            return result;
        } finally {
            workbook.close();
        }
    }

    public static List<Map<Integer, String>> readMaps(InputStream inputStream, ReadOptions options)
        throws IOException {
        Workbook workbook = new XSSFWorkbook(inputStream);
        try {
            Sheet sheet = workbook.getSheetAt(options.getSheetIndex());
            DataFormatter formatter = new DataFormatter();
            List<Map<Integer, String>> result = new ArrayList<Map<Integer, String>>();
            for (int rowIndex = options.getHeadRowNumber(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null || rowIsEmpty(row)) {
                    continue;
                }
                Map<Integer, String> map = new LinkedHashMap<Integer, String>();
                for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    map.put(cellIndex, formatter.formatCellValue(cell));
                }
                result.add(map);
            }
            return result;
        } finally {
            workbook.close();
        }
    }

    public static <T> void fill(InputStream template, OutputStream outputStream, List<T> data,
        FillOptions options) throws IOException {
        Workbook workbook = new XSSFWorkbook(template);
        try {
            Sheet sheet = workbook.getSheetAt(0);
            replaceScalarCells(workbook, sheet, options.getScalars());
            fillListRows(workbook, sheet, data);
            workbook.write(outputStream);
        } finally {
            workbook.close();
        }
    }

    private static int writeHead(Workbook workbook, Sheet sheet, Class<?> headClass, List<ColumnMeta> columns) {
        int headRows = 1;
        for (ColumnMeta column : columns) {
            headRows = Math.max(headRows, column.heads.length);
        }
        ExcelRowHeight rowHeight = headClass.getAnnotation(ExcelRowHeight.class);
        for (int rowIndex = 0; rowIndex < headRows; rowIndex++) {
            Row row = sheet.createRow(rowIndex);
            if (rowHeight != null && rowHeight.head() > -1) {
                row.setHeightInPoints(rowHeight.head());
            }
            for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
                ColumnMeta column = columns.get(columnIndex);
                Cell cell = row.createCell(columnIndex, CellType.STRING);
                String head = column.headAt(rowIndex);
                if (columnIndex > 0 && head.length() > 0 && head.equals(columns.get(columnIndex - 1).headAt(rowIndex))) {
                    head = "";
                }
                cell.setCellValue(head);
            }
        }
        mergeHead(sheet, columns, headRows);
        applyColumnWidths(sheet, headClass, columns);
        return headRows;
    }

    private static void writeRows(Workbook workbook, Sheet sheet, List<ColumnMeta> columns, List<?> data,
        int startRow) {
        ExcelRowHeight rowHeight = null;
        if (!columns.isEmpty()) {
            rowHeight = columns.get(0).field.getDeclaringClass().getAnnotation(ExcelRowHeight.class);
        }
        Map<String, CellStyle> styles = new LinkedHashMap<String, CellStyle>();
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(startRow + i);
            if (rowHeight != null && rowHeight.content() > -1) {
                row.setHeightInPoints(rowHeight.content());
            }
            Object bean = data.get(i);
            for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
                ColumnMeta column = columns.get(columnIndex);
                Cell cell = row.createCell(columnIndex);
                Object value = fieldValue(column.field, bean);
                if (column.loopMerge != null && i % column.loopMerge.eachRow() != 0) {
                    cell.setBlank();
                    continue;
                }
                writeValue(workbook, sheet, cell, value, column, styles);
            }
        }
        applyLoopMerges(sheet, columns, data.size(), startRow);
    }

    private static void writeValue(Workbook workbook, Sheet sheet, Cell cell, Object value, ColumnMeta column,
        Map<String, CellStyle> styles) {
        if (value == null) {
            cell.setBlank();
            return;
        }
        if (value instanceof WriteCellData) {
            WriteCellData writeCellData = (WriteCellData)value;
            cell.setCellValue(writeCellData.getText());
            for (ImageCellData image : writeCellData.getImages()) {
                addPicture(workbook, sheet, image.getImage(), cell.getRowIndex() + image.getFirstRowOffset(),
                    cell.getColumnIndex() + image.getFirstColumnOffset(),
                    cell.getRowIndex() + image.getLastRowOffset() + 1,
                    cell.getColumnIndex() + image.getLastColumnOffset() + 1);
            }
            return;
        }
        byte[] image = imageBytes(value);
        if (image != null) {
            cell.setCellValue("");
            addPicture(workbook, sheet, image, cell.getRowIndex(), cell.getColumnIndex(),
                cell.getRowIndex() + 1, cell.getColumnIndex() + 1);
            return;
        }
        if (value instanceof Number) {
            cell.setCellValue(((Number)value).doubleValue());
            applyNumberStyle(workbook, cell, column, styles);
            return;
        }
        if (value instanceof LocalDateTime) {
            LocalDateTime time = (LocalDateTime)value;
            Date date = Date.from(time.atZone(ZoneId.systemDefault()).toInstant());
            cell.setCellValue(date);
            applyDateStyle(workbook, cell, column, styles);
            return;
        }
        if (value instanceof Date) {
            cell.setCellValue((Date)value);
            applyDateStyle(workbook, cell, column, styles);
            return;
        }
        cell.setCellValue(String.valueOf(value));
    }

    private static void applyDateStyle(Workbook workbook, Cell cell, ColumnMeta column, Map<String, CellStyle> styles) {
        String pattern = column.datePattern == null ? "yyyy-MM-dd HH:mm:ss" : column.datePattern.value();
        cell.setCellStyle(styleFor(workbook, styles, "date:" + pattern, pattern));
    }

    private static void applyNumberStyle(Workbook workbook, Cell cell, ColumnMeta column, Map<String, CellStyle> styles) {
        if (column.numberPattern == null) {
            return;
        }
        String pattern = column.numberPattern.value();
        cell.setCellStyle(styleFor(workbook, styles, "number:" + pattern, pattern));
    }

    private static CellStyle styleFor(Workbook workbook, Map<String, CellStyle> styles, String key, String format) {
        CellStyle style = styles.get(key);
        if (style != null) {
            return style;
        }
        style = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();
        style.setDataFormat(dataFormat.getFormat(format));
        styles.put(key, style);
        return style;
    }

    private static void mergeHead(Sheet sheet, List<ColumnMeta> columns, int headRows) {
        for (int rowIndex = 0; rowIndex < headRows; rowIndex++) {
            int start = -1;
            String current = null;
            for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
                String value = columns.get(columnIndex).headAt(rowIndex);
                if (value.length() == 0) {
                    continue;
                }
                if (current == null || !current.equals(value)) {
                    if (start >= 0 && columnIndex - start > 1) {
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, start, columnIndex - 1));
                    }
                    current = value;
                    start = columnIndex;
                }
            }
            if (start >= 0 && columns.size() - start > 1) {
                sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, start, columns.size() - 1));
            }
        }
    }

    private static void applyLoopMerges(Sheet sheet, List<ColumnMeta> columns, int rows, int startRow) {
        for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
            ColumnMeta column = columns.get(columnIndex);
            if (column.loopMerge == null) {
                continue;
            }
            int each = column.loopMerge.eachRow();
            for (int i = 0; i < rows; i += each) {
                int end = Math.min(i + each - 1, rows - 1);
                if (end > i) {
                    sheet.addMergedRegion(new CellRangeAddress(startRow + i, startRow + end, columnIndex, columnIndex));
                }
            }
        }
    }

    private static void applyColumnWidths(Sheet sheet, Class<?> headClass, List<ColumnMeta> columns) {
        ExcelColumnWidth classWidth = headClass.getAnnotation(ExcelColumnWidth.class);
        for (int i = 0; i < columns.size(); i++) {
            ExcelColumnWidth width = columns.get(i).field.getAnnotation(ExcelColumnWidth.class);
            int value = width == null ? (classWidth == null ? -1 : classWidth.value()) : width.value();
            if (value > -1) {
                sheet.setColumnWidth(i, value * 256);
            }
        }
    }

    private static <T> T readBean(Row row, List<ColumnMeta> columns, Map<String, Integer> headerIndex,
        Class<T> headClass) {
        try {
            Constructor<T> constructor = headClass.getDeclaredConstructor();
            constructor.setAccessible(true);
            T bean = constructor.newInstance();
            for (int columnOrder = 0; columnOrder < columns.size(); columnOrder++) {
                ColumnMeta column = columns.get(columnOrder);
                Integer cellIndex = headerIndex.get(column.primaryHead());
                if (cellIndex == null) {
                    cellIndex = columnOrder;
                }
                Cell cell = row.getCell(cellIndex);
                column.field.setAccessible(true);
                column.field.set(bean, convertCell(cell, column.field.getType()));
            }
            return bean;
        } catch (Exception e) {
            throw new IllegalStateException("Cannot read row " + row.getRowNum(), e);
        }
    }

    private static Object convertCell(Cell cell, Class<?> targetType) {
        if (cell == null) {
            return null;
        }
        DataFormatter formatter = new DataFormatter();
        if (targetType == String.class) {
            return formatter.formatCellValue(cell);
        }
        if (targetType == Double.class || targetType == double.class) {
            if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getNumericCellValue();
            }
            String text = formatter.formatCellValue(cell);
            return text.length() == 0 ? null : Double.valueOf(text);
        }
        if (targetType == LocalDateTime.class) {
            if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getLocalDateTimeCellValue().withNano(0);
            }
            String text = formatter.formatCellValue(cell);
            if (text.length() == 0) {
                return null;
            }
            return LocalDateTime.parse(text, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        }
        return formatter.formatCellValue(cell);
    }

    private static Map<String, Integer> headerIndex(Sheet sheet, int headRowNumber) {
        Map<String, Integer> result = new LinkedHashMap<String, Integer>();
        if (headRowNumber <= 0) {
            return result;
        }
        Row row = sheet.getRow(headRowNumber - 1);
        if (row == null) {
            return result;
        }
        DataFormatter formatter = new DataFormatter();
        for (int i = 0; i < row.getLastCellNum(); i++) {
            String value = formatter.formatCellValue(row.getCell(i));
            if (value.length() > 0) {
                result.put(value, i);
            }
        }
        return result;
    }

    private static void replaceScalarCells(Workbook workbook, Sheet sheet, Map<String, Object> scalars) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String value = cell.getStringCellValue();
                    if (value.startsWith("{") && value.endsWith("}") && !value.startsWith("{.")) {
                        String key = value.substring(1, value.length() - 1);
                        writeFillValue(workbook, cell, scalars.get(key), null);
                    }
                }
            }
        }
    }

    private static <T> void fillListRows(Workbook workbook, Sheet sheet, List<T> data) {
        int templateRowIndex = findListTemplateRow(sheet);
        if (templateRowIndex < 0 || data.isEmpty()) {
            return;
        }
        Row templateRow = sheet.getRow(templateRowIndex);
        List<String> placeholders = new ArrayList<String>();
        for (int i = 0; i < templateRow.getLastCellNum(); i++) {
            Cell cell = templateRow.getCell(i);
            String text = cell == null ? "" : cell.getStringCellValue();
            placeholders.add(text.startsWith("{.") && text.endsWith("}") ? text.substring(2, text.length() - 1) : "");
        }
        if (data.size() > 1 && templateRowIndex < sheet.getLastRowNum()) {
            sheet.shiftRows(templateRowIndex + 1, sheet.getLastRowNum(), data.size() - 1, true, false);
        }
        for (int rowOffset = 0; rowOffset < data.size(); rowOffset++) {
            Row row = sheet.getRow(templateRowIndex + rowOffset);
            if (row == null) {
                row = sheet.createRow(templateRowIndex + rowOffset);
            }
            Object bean = data.get(rowOffset);
            for (int columnIndex = 0; columnIndex < placeholders.size(); columnIndex++) {
                Cell cell = row.getCell(columnIndex);
                if (cell == null) {
                    cell = row.createCell(columnIndex);
                }
                String fieldName = placeholders.get(columnIndex);
                if (fieldName.length() > 0) {
                    Field field = findField(bean.getClass(), fieldName);
                    writeFillValue(workbook, cell, fieldValue(field, bean), field);
                }
            }
        }
    }

    private static int findListTemplateRow(Sheet sheet) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().startsWith("{.")) {
                    return row.getRowNum();
                }
            }
        }
        return -1;
    }

    private static void writeFillValue(Workbook workbook, Cell cell, Object value, Field field) {
        if (value == null) {
            cell.setBlank();
            return;
        }
        if (value instanceof Number) {
            cell.setCellValue(((Number)value).doubleValue());
            return;
        }
        if (value instanceof LocalDateTime) {
            LocalDateTime time = (LocalDateTime)value;
            cell.setCellValue(Date.from(time.atZone(ZoneId.systemDefault()).toInstant()));
            String pattern = "yyyy-MM-dd HH:mm:ss";
            if (field != null && field.getAnnotation(ExcelDatePattern.class) != null) {
                pattern = field.getAnnotation(ExcelDatePattern.class).value();
            }
            CellStyle style = workbook.createCellStyle();
            style.setDataFormat(workbook.createDataFormat().getFormat(pattern));
            cell.setCellStyle(style);
            return;
        }
        cell.setCellValue(String.valueOf(value));
    }

    private static List<ColumnMeta> columnsForWrite(Class<?> headClass, WriteOptions options) {
        List<ColumnMeta> all = columnsForRead(headClass);
        List<ColumnMeta> filtered = new ArrayList<ColumnMeta>();
        for (ColumnMeta column : all) {
            if (!options.getIncludeFields().isEmpty() && !options.getIncludeFields().contains(column.field.getName())) {
                continue;
            }
            if (options.getExcludeFields().contains(column.field.getName())) {
                continue;
            }
            filtered.add(column);
        }
        return filtered;
    }

    private static List<ColumnMeta> columnsForRead(Class<?> headClass) {
        List<ColumnMeta> columns = new ArrayList<ColumnMeta>();
        Field[] fields = headClass.getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (field.getAnnotation(ExcelIgnoreColumn.class) != null) {
                continue;
            }
            columns.add(new ColumnMeta(field, i));
        }
        Collections.sort(columns, new Comparator<ColumnMeta>() {
            @Override
            public int compare(ColumnMeta left, ColumnMeta right) {
                int leftIndex = left.index < 0 ? Integer.MAX_VALUE : left.index;
                int rightIndex = right.index < 0 ? Integer.MAX_VALUE : right.index;
                if (leftIndex != rightIndex) {
                    return leftIndex - rightIndex;
                }
                return left.order - right.order;
            }
        });
        return columns;
    }

    private static boolean rowIsEmpty(Row row) {
        DataFormatter formatter = new DataFormatter();
        for (int i = 0; i < row.getLastCellNum(); i++) {
            if (formatter.formatCellValue(row.getCell(i)).length() > 0) {
                return false;
            }
        }
        return true;
    }

    private static Field findField(Class<?> type, String placeholder) {
        for (Field field : type.getDeclaredFields()) {
            ExcelColumn column = field.getAnnotation(ExcelColumn.class);
            if (field.getName().equals(placeholder)) {
                return field;
            }
            if (column != null) {
                for (String head : column.value()) {
                    if (head.equals(placeholder)) {
                        return field;
                    }
                }
            }
        }
        throw new IllegalArgumentException("No field for placeholder: " + placeholder);
    }

    private static Object fieldValue(Field field, Object bean) {
        try {
            field.setAccessible(true);
            return field.get(bean);
        } catch (Exception e) {
            throw new IllegalStateException("Cannot read field " + field.getName(), e);
        }
    }

    private static byte[] imageBytes(Object value) {
        try {
            if (value instanceof byte[]) {
                return ((byte[])value).clone();
            }
            if (value instanceof File) {
                return readAll(new FileInputStream((File)value));
            }
            if (value instanceof InputStream) {
                return readAll((InputStream)value);
            }
            if (value instanceof URL) {
                return readAll(((URL)value).openStream());
            }
            if (value instanceof String) {
                File file = new File((String)value);
                if (file.exists()) {
                    return readAll(new FileInputStream(file));
                }
            }
            return null;
        } catch (IOException e) {
            throw new IllegalStateException("Cannot load image", e);
        }
    }

    private static byte[] readAll(InputStream inputStream) throws IOException {
        try {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            byte[] buffer = new byte[4096];
            int len;
            while ((len = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, len);
            }
            return outputStream.toByteArray();
        } finally {
            inputStream.close();
        }
    }

    private static void addPicture(Workbook workbook, Sheet sheet, byte[] image, int firstRow, int firstColumn,
        int lastRow, int lastColumn) {
        int pictureIndex = workbook.addPicture(image, Workbook.PICTURE_TYPE_JPEG);
        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setRow1(firstRow);
        anchor.setCol1(firstColumn);
        anchor.setRow2(lastRow);
        anchor.setCol2(lastColumn);
        drawing.createPicture(anchor, pictureIndex);
    }

    private static final class ColumnMeta {
        private final Field field;
        private final int order;
        private final int index;
        private final String[] heads;
        private final ExcelDatePattern datePattern;
        private final ExcelNumberPattern numberPattern;
        private final ExcelLoopMerge loopMerge;

        private ColumnMeta(Field field, int order) {
            this.field = field;
            this.order = order;
            ExcelColumn column = field.getAnnotation(ExcelColumn.class);
            this.index = column == null ? -1 : column.index();
            this.heads = column == null || column.value().length == 0
                ? new String[] {field.getName()} : column.value();
            this.datePattern = field.getAnnotation(ExcelDatePattern.class);
            this.numberPattern = field.getAnnotation(ExcelNumberPattern.class);
            this.loopMerge = field.getAnnotation(ExcelLoopMerge.class);
        }

        private String headAt(int rowIndex) {
            if (heads.length == 0) {
                return "";
            }
            if (rowIndex < heads.length) {
                return heads[rowIndex];
            }
            return "";
        }

        private String primaryHead() {
            return heads.length == 0 ? field.getName() : heads[heads.length - 1];
        }
    }
}
