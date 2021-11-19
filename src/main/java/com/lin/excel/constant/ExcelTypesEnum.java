package com.lin.excel.constant;

/**
 * excel 类型枚举类
 * xls 为2003版
 * xlsx 为2007版
 */
public enum ExcelTypesEnum {

    xls("xls", "应用模块"),
    xlsx("xlsx", "模块");


    ExcelTypesEnum(String code, String name) {
        this.code = code;
        this.name = name;
    }

    private String code;
    private String name;

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}