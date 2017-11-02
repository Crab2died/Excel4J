/*
 *
 *                  Copyright 2017 Crab2Died
 *                     All rights reserved.
 *
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * Browse for more information ：
 * 1) https://gitee.com/Crab2Died/Excel4J
 * 2) https://github.com/Crab2died/Excel4J
 *
 */

package com.github.crab2died.utils;

import com.github.crab2died.annotation.ExcelField;
import com.github.crab2died.converter.DefaultConvertible;
import com.github.crab2died.converter.WriteConvertible;
import com.github.crab2died.handler.ExcelHeader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.*;
import java.util.regex.Pattern;


public class Utils {

    /**
     * <p>getter与setter方法的枚举</p>
     *
     * @author Crab2Died
     */
    private enum MethodType {

        GET("get"), SET("set");

        private String value;

        MethodType(String value) {
            this.value = value;
        }

        public String getValue() {
            return value;
        }
    }

    /**
     * <p>根据JAVA对象注解获取Excel表头信息</p>
     *
     * @param clz 类型
     * @return 表头信息
     * @throws IllegalAccessException 异常
     * @throws InstantiationException 异常
     */
    static
    public List<ExcelHeader> getHeaderList(Class<?> clz) throws IllegalAccessException,
            InstantiationException {
        List<ExcelHeader> headers = new ArrayList<>();
        List<Field> fields = new ArrayList<>();
        for (Class<?> clazz = clz; clazz != Object.class; clazz = clazz.getSuperclass()) {
            fields.addAll(Arrays.asList(clazz.getDeclaredFields()));
        }
        for (Field field : fields) {
            // 是否使用ExcelField注解
            if (field.isAnnotationPresent(ExcelField.class)) {
                ExcelField er = field.getAnnotation(ExcelField.class);
                headers.add(new ExcelHeader(er.title(), er.order(), er.writeConverter().newInstance(),
                        er.readConverter().newInstance(), field.getName(), field.getType()));
            }
        }
        Collections.sort(headers);
        return headers;
    }

    /**
     * 获取excel列表头
     *
     * @param titleRow excel行
     * @param clz      类型
     * @return ExcelHeader集合
     * @throws InstantiationException 异常
     * @throws IllegalAccessException 异常
     */
    static
    public Map<Integer, ExcelHeader> getHeaderMap(Row titleRow, Class<?> clz)
            throws InstantiationException, IllegalAccessException {

        List<ExcelHeader> headers = getHeaderList(clz);
        Map<Integer, ExcelHeader> maps = new HashMap<>();
        for (Cell c : titleRow) {
            String title = c.getStringCellValue();
            for (ExcelHeader eh : headers) {
                if (eh.getTitle().equals(title.trim())) {
                    maps.put(c.getColumnIndex(), eh);
                    break;
                }
            }
        }
        return maps;
    }

    /**
     * 获取单元格内容
     *
     * @param c 单元格
     * @return 单元格内容
     */
    static
    public String getCellValue(Cell c) {
        String o;
        switch (c.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                o = "";
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                o = String.valueOf(c.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                o = String.valueOf(c.getCellFormula());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                o = String.valueOf(c.getNumericCellValue());
                o = matchDoneBigDecimal(o);
                o = RegularUtils.converNumByReg(o);
                break;
            case Cell.CELL_TYPE_STRING:
                o = c.getStringCellValue();
                break;
            default:
                o = null;
                break;
        }
        return o;
    }

    /**
     * 字符串转对象
     *
     * @param strField 字符串
     * @param clazz    待转类型
     * @return 转换后数据
     */
    static
    public Object str2TargetClass(String strField, Class<?> clazz) {
        if (null == strField || "".equals(strField))
            return null;
        if ((Long.class == clazz) || (long.class == clazz)) {
            strField = matchDoneBigDecimal(strField);
            strField = RegularUtils.converNumByReg(strField);
            return Long.parseLong(strField);
        }
        if ((Integer.class == clazz) || (int.class == clazz)) {
            strField = matchDoneBigDecimal(strField);
            strField = RegularUtils.converNumByReg(strField);
            return Integer.parseInt(strField);
        }
        if ((Float.class == clazz) || (float.class == clazz)) {
            strField = matchDoneBigDecimal(strField);
            return Float.parseFloat(strField);
        }
        if ((Double.class == clazz) || (double.class == clazz)) {
            strField = matchDoneBigDecimal(strField);
            return Double.parseDouble(strField);
        }
        if ((Character.class == clazz) || (char.class == clazz)) {
            return strField.toCharArray()[0];
        }
        if ((Boolean.class == clazz) || (boolean.class == clazz)) {
            return Boolean.parseBoolean(strField);
        }
        if (Date.class == clazz) {
            return DateUtils.str2DateUnmatch2Null(strField);
        }
        return strField;
    }

    /**
     * 科学计数法数据转换
     *
     * @param bigDecimal 科学计数法
     * @return 数据字符串
     */
    private static String matchDoneBigDecimal(String bigDecimal) {
        // 对科学计数法进行处理
        boolean flg = Pattern.matches("^-?\\d+(\\.\\d+)?(E-?\\d+)?$", bigDecimal);
        if (flg) {
            BigDecimal bd = new BigDecimal(bigDecimal);
            bigDecimal = bd.toPlainString();
        }
        return bigDecimal;
    }

    /**
     * 获取字段的getter或setter方法
     *
     * @param fieldClass 字段类型
     * @param fieldName  字段名
     * @param methodType getter或setter
     * @return getter或setter方法
     */
    private static String getOrSet(Class fieldClass, String fieldName, MethodType methodType) {

        if (null == fieldClass || null == fieldName)
            return null;

        // 对boolean类型的特殊处理
        if (boolean.class == fieldClass) {
            if (MethodType.SET == methodType) {
                if (fieldName.startsWith("is") &&
                        Character.isUpperCase(fieldName.substring(2, 3).toCharArray()[0])) {
                    return methodType.getValue() + fieldName.substring(2);
                }
            }
            if (MethodType.GET == methodType) {
                if (fieldName.startsWith("is") &&
                        Character.isUpperCase(fieldName.substring(2, 3).toCharArray()[0])) {
                    return fieldName;
                } else {
                    return "is" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                }
            }
        }
        // 对Boolean类型的特殊处理
        if (Boolean.class == fieldClass) {
            if (MethodType.SET == methodType) {
                if (fieldName.startsWith("is") &&
                        Character.isUpperCase(fieldName.substring(2, 3).toCharArray()[0])) {
                    return methodType.getValue() + fieldName.substring(2);
                }
            }
            if (MethodType.GET == methodType) {
                if (fieldName.startsWith("is") &&
                        Character.isUpperCase(fieldName.substring(2, 3).toCharArray()[0])) {
                    return methodType.getValue() + fieldName.substring(2);
                }
            }
        }
        return methodType.getValue() + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
    }

    /**
     * <p>根据对象的属性名{@code fieldName}获取某个java的属性{@link java.lang.reflect.Field}</p>
     *
     * @param clazz     java对象的class属性
     * @param fieldName 属性名
     * @return {@link java.lang.reflect.Field}   java对象的属性
     * @author Crab2Died
     */
    private static Field matchClassField(Class clazz, String fieldName) {
        List<Field> fields = new ArrayList<>();
        for (; clazz != Object.class; clazz = clazz.getSuperclass()) {
            fields.addAll(Arrays.asList(clazz.getDeclaredFields()));
        }
        for (Field field : fields) {
            if (fieldName.equals(field.getName())) {
                return field;
            }
        }
        return null;
    }

    /**
     * 根据属性名与属性类型获取字段内容
     *
     * @param bean             对象
     * @param fieldName        字段名
     * @param fieldClass       字段类型
     * @param writeConvertible 写入转换器
     * @return 对象指定字段内容
     * @throws NoSuchMethodException     异常
     * @throws InvocationTargetException 异常
     * @throws IllegalAccessException    异常
     */
    static
    public String getProperty(Object bean, String fieldName, Class fieldClass, WriteConvertible writeConvertible)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        Method method = bean.getClass().getDeclaredMethod(getOrSet(fieldClass, fieldName, MethodType.GET));
        Object object = method.invoke(bean);
        if (null != writeConvertible && writeConvertible.getClass() != DefaultConvertible.class) {
            // 写入转换器
            object = writeConvertible.execWrite(object);
        }
        return object.toString();
    }

    static
    public void copyProperty(Object bean, String name, Object value) throws NoSuchMethodException,
            InvocationTargetException, IllegalAccessException {
        Field field = matchClassField(bean.getClass(), name);
        if (null == field)
            return;
        Method method = bean.getClass().getDeclaredMethod(
                getOrSet(field.getType(), name, MethodType.SET),
                field.getType()
        );
        if (value.getClass() == field.getType()) {
            method.invoke(bean, value);
        } else {
            method.invoke(bean, str2TargetClass(value.toString(), field.getType()));
        }
    }
}
