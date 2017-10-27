/*
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
     * <p>根据JAVA对象注解获取Excel表头信息</p>
     *
     * @param clz 类型
     * @return 表头信息
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
        if (Date.class == clazz) {
            return DateUtils.str2DateUnmatch2Null(strField);
        }
        return strField;
    }

    private static String matchDoneBigDecimal(String bigDecimal) {
        // 对科学计数法进行处理
        boolean flg = Pattern.matches("^-?\\d+(\\.\\d+)?(E-?\\d+)?$", bigDecimal);
        if (flg) {
            BigDecimal bd = new BigDecimal(bigDecimal);
            bigDecimal = bd.toPlainString();
        }
        return bigDecimal;
    }

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

    static
    public String getProperty(Object bean, String fieldName, Class fieldClass, WriteConvertible writeConvertible)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException, InstantiationException {

        Method method = bean.getClass().getDeclaredMethod(getOrSet(fieldClass, fieldName, MethodType.GET));
        Object object = method.invoke(bean);
        if (null != writeConvertible && writeConvertible.getClass() != DefaultConvertible.class) {
            object = writeConvertible.execWrite(object);
        }
        return object.toString();
    }

}
