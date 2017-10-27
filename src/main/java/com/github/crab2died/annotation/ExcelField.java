package com.github.crab2died.annotation;

import com.github.crab2died.converter.DefaultConvertible;
import com.github.crab2died.converter.ReadConvertible;
import com.github.crab2died.converter.WriteConvertible;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

/**
 * 功能说明: 用来在对象的属性上加入的annotation，通过该annotation说明某个属性所对应的标题
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelField {

    /*
     * 属性的标题名称
     */
    String title();

    /*
     * 写数据转换器
     */
    Class<? extends WriteConvertible> writeConverter()
            default DefaultConvertible.class;

    /*
     * 读数据转换器
     */
    Class<? extends ReadConvertible> readConverter()
            default DefaultConvertible.class;

    /*
     * 在excel的顺序
     */
    int order() default 9999;
}
