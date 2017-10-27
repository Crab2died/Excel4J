package com.github.crab2died.converter;

/**
 * 默认转换器
 */
public class DefaultConvertible implements WriteConvertible, ReadConvertible {

    @Override
    public Object execWrite(Object object) {
        return object;
    }

    @Override
    public Object execRead(String object) {
        return object;
    }
}
