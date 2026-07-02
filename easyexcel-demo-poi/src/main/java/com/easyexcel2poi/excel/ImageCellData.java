package com.easyexcel2poi.excel;

public final class ImageCellData {
    private final byte[] image;
    private final int firstRowOffset;
    private final int firstColumnOffset;
    private final int lastRowOffset;
    private final int lastColumnOffset;

    public ImageCellData(byte[] image, int firstRowOffset, int firstColumnOffset, int lastRowOffset,
        int lastColumnOffset) {
        this.image = image.clone();
        this.firstRowOffset = firstRowOffset;
        this.firstColumnOffset = firstColumnOffset;
        this.lastRowOffset = lastRowOffset;
        this.lastColumnOffset = lastColumnOffset;
    }

    public byte[] getImage() {
        return image.clone();
    }

    public int getFirstRowOffset() {
        return firstRowOffset;
    }

    public int getFirstColumnOffset() {
        return firstColumnOffset;
    }

    public int getLastRowOffset() {
        return lastRowOffset;
    }

    public int getLastColumnOffset() {
        return lastColumnOffset;
    }
}
