package com.github.crab2died.utils;

/**
 * <p>getter与setter方法的枚举</p>
 * @author  Crab2Died
 */
public enum MethodType {

    GET("get"), SET("set");

    private String value;

    MethodType(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}
