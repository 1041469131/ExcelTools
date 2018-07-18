package com.hqjl.common.excel.bean;

import java.util.List;

/**
 * ExcelSheetData class
 *
 * @author LiXiang
 * @date 2018/01/22
 */
public class ExcelSheetData {
    private String sheetName;
    private List<ExcelTableData> tables;

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<ExcelTableData> getTables() {
        return tables;
    }

    public void setTables(List<ExcelTableData> tables) {
        this.tables = tables;
    }
}
