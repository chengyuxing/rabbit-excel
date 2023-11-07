package com.github.chengyuxing.excel.type;

/**
 * Sheet Cell Coord helper
 */
public class Coord {
    private final int x;
    private final int y;

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
