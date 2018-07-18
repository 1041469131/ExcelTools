package com.hqjl.common.excel.bean;

import java.util.List;
import java.util.Map;

/**
 * ExcelTableData class
 *
 * @author LiXiang
 * @date 2018/01/22
 */
public class ExcelTableData {
    private int startRow;
    private int startColumn;
    private ExcelCellData title;
    private List<List<ExcelCellData>> headers;
    private List<Map<String,Object>> rows;
    private int numType;

    public ExcelCellData getTitle() {
        return title;
    }

    public void setTitle(ExcelCellData title) {
        this.title = title;
    }

    public List<List<ExcelCellData>> getHeaders() {
        return headers;
    }

    public void setHeaders(List<List<ExcelCellData>> headers) {
        this.headers = headers;
    }

    public List<Map<String, Object>> getRows() {
        return rows;
    }

    public void setRows(List<Map<String, Object>> rows) {
        this.rows = rows;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getStartColumn() {
        return startColumn;
    }

    public void setStartColumn(int startColumn) {
        this.startColumn = startColumn;
    }

    public int getNumType() {
        return numType;
    }

    public void setNumType(int numType) {
        this.numType = numType;
    }
}
