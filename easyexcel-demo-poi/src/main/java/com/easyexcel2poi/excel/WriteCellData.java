package com.easyexcel2poi.excel;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public final class WriteCellData {
    private final String text;
    private final List<ImageCellData> images = new ArrayList<ImageCellData>();

    public WriteCellData(String text) {
        this.text = text;
    }

    public String getText() {
        return text;
    }

    public List<ImageCellData> getImages() {
        return Collections.unmodifiableList(images);
    }

    public void addImage(ImageCellData image) {
        images.add(image);
    }
}
