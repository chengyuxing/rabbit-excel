package com.github.chengyuxing.excel.type;

/**
 * 读取到的Sheet元数据
 */
public class SheetInfo {
    private final int index;
    private final String name;
    private final int size;

    private SheetInfo(int index, String name, int size) {
        this.index = index;
        this.name = name;
        this.size = size;
    }

    public static SheetInfo of(int index, String name, int size) {
        return new SheetInfo(index, name, size);
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