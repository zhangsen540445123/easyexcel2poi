package com.easyexcel2poi.excel;

import java.util.ArrayList;
import java.util.List;

public abstract class PageReadListener<T> implements ReadListener<T> {
    private final int pageSize;
    private final List<T> buffer = new ArrayList<T>();

    protected PageReadListener(int pageSize) {
        this.pageSize = pageSize;
    }

    @Override
    public final void invoke(T data) {
        buffer.add(data);
        if (buffer.size() >= pageSize) {
            flush();
        }
    }

    @Override
    public final void doAfterAllAnalysed() {
        flush();
    }

    protected abstract void invokePage(List<T> data);

    private void flush() {
        if (buffer.isEmpty()) {
            return;
        }
        invokePage(new ArrayList<T>(buffer));
        buffer.clear();
    }
}
