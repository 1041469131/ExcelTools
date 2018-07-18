package com.hqjl.common.excel;

import com.hqjl.common.excel.bean.CellBean;

import java.util.List;
import java.util.Map;

/**
 * ExcelTableData class
 *
 * @author LiXiang
 * @date 2018/01/22
 */
public class ExcelTableData {
    private ExcelCellData title;
    private List<List<ExcelCellData>> headers;
    private List<Map<String,Object>> rows;

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
}
