package com.github.chengyuxing.excel.core;

public class SheetMetaData {
    private int index;
    private String name;
    private int size;

    private SheetMetaData(int index, String name, int size) {
        this.index = index;
        this.name = name;
        this.size = size;
    }

    static SheetMetaData of(int index, String name, int size) {
        return new SheetMetaData(index, name, size);
    }

    public int getIndex() {
        return index;
    }

    public String getName() {
        return name;
    }

    public int getSize() {
        return size;
    }

    @Override
    public String toString() {
        return "SheetMetaData{" +
                "index=" + index +
                ", name='" + name + '\'' +
                ", size=" + size +
                '}';
    }
}