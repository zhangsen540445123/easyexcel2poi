package com.easyexcel2poi.excel;

public interface ReadListener<T> {
    void invoke(T data);

    void doAfterAllAnalysed();
}
