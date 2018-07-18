package com.hqjl.common.excel.bean;

/**
 * @author xiang.li
 * create date 2018/7/12
 * description
 */
public enum FormatType {
    /**
     * 文本
     */
    TXT(0),
    /**
     * 小数
     */
    DECIMALS(1),
    /**
     * 百分比
     */
    PERCENT(8),

    /**
     * 字符窜
     */
    STR(2);
    private int numType;

    FormatType(int numType){
        this.numType=numType;
    }

    public int getNumType() {
        return numType;
    }
    public static FormatType valueOf(int numType) {
        for (FormatType classType : FormatType.values()) {
            if (classType.numType == numType) {
                return classType;
            }
        }
        return null;
    }
}
