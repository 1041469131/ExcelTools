package com.hqjl.common.excel;

/**
 * ExcelCellData class
 *
 * @author LiXiang
 * @date 2018/01/22
 */
public class ExcelCellData {
    private String key;
    private String name;
    private int xSize;
    /**纵向合并个数*/
    private int ySize;

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getxSize() {
        return xSize;
    }

    public void setxSize(int xSize) {
        this.xSize = xSize;
    }

    public int getySize() {
        return ySize;
    }

    public void setySize(int ySize) {
        this.ySize = ySize;
    }

    public ExcelCellData() {
    }

    public ExcelCellData(String key, String name, int xSize, int ySize) {
        this.key = key;
        this.name = name;
        this.xSize = xSize;
        this.ySize = ySize;
    }

    public ExcelCellData(String name, int xSize, int ySize) {
        this.name = name;
        this.xSize = xSize;
        this.ySize = ySize;
    }
}
