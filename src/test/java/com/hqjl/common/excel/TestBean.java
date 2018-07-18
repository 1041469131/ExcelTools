package com.hqjl.common.excel;

import com.hqjl.common.excel.annotation.ExcelVoConfig;
import com.hqjl.common.excel.annotation.input.InputDateConfig;
import com.hqjl.common.excel.bean.BaseExcelVo;

import java.util.Date;
@ExcelVoConfig
public class TestBean extends BaseExcelVo {

    private int id;
    private String name;
    @InputDateConfig
    private Date date;

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    @Override
    public int getHashVal() {
        return 0;
    }
}
