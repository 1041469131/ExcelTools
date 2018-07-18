package com.hqjl.common.excel;

import com.hqjl.common.excel.annotation.ExcelVoConfig;
import com.hqjl.common.excel.bean.BaseExcelVo;

/**
 * TestMode class
 *
 * @author LiXiang
 * @date 2018/01/18
 */
@ExcelVoConfig
public class TestMode extends BaseExcelVo {
    private Integer a;
    private Integer b;
    private Integer c;
    private Integer d;
    private Integer e;

    public TestMode(Integer a, Integer b, Integer c, Integer d, Integer e) {
        this.a = a;
        this.b = b;
        this.c = c;
        this.d = d;
        this.e = e;
    }

    public Integer getA() {
        return a;
    }

    public void setA(Integer a) {
        this.a = a;
    }

    public Integer getB() {
        return b;
    }

    public void setB(Integer b) {
        this.b = b;
    }

    public Integer getC() {
        return c;
    }

    public void setC(Integer c) {
        this.c = c;
    }

    public Integer getD() {
        return d;
    }

    public void setD(Integer d) {
        this.d = d;
    }

    public Integer getE() {
        return e;
    }

    public void setE(Integer e) {
        this.e = e;
    }

    @Override
    public int getHashVal() {
        return 0;
    }
}
