package com.github.chengyuxing.excel.type;

/**
 * 正整数坐标帮助类
 */
public class Coord {
    private final int x;
    private final int y;

    /**
     * 构造函数
     *
     * @param x 横坐标
     * @param y 纵坐标
     */
    public Coord(int x, int y) {
        if (x < 0 || y < 0) {
            throw new IllegalArgumentException("x and y must not be negative.");
        }
        this.x = x;
        this.y = y;
    }

    public int getX() {
        return x;
    }

    public int getY() {
        return y;
    }

    @Override
    public String toString() {
        return "Coord{" +
                "x=" + x +
                ", y=" + y +
                '}';
    }
}
