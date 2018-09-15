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

package com.github.crab2died.handler;

import com.github.crab2died.converter.ReadConvertible;
import com.github.crab2died.converter.WriteConvertible;
import lombok.Data;

/**
 * 功能说明: 用来存储Excel标题的对象，通过该对象可以获取标题和方法的对应关系
 * @author junmingyang
 */
@Data
public class ExcelHeader implements Comparable<ExcelHeader> {

    /**
     * excel的标题名称
     */
    private String title;

    /**
     * 每一个标题的顺序
     */
    private int order;

    /**
     * 写数据转换器
     */
    private WriteConvertible writeConverter;

    /**
     * 读数据转换器
     */
    private ReadConvertible readConverter;

    /**
     * 注解域
     */
    private String filed;

    /**
     * 属性类型
     */
    private Class<?> filedClazz;


    @Override
    public int compareTo(ExcelHeader o) {
        return order - o.order;
    }

    public ExcelHeader() {
        super();
    }

    public ExcelHeader(String title, int order, WriteConvertible writeConverter,
                       ReadConvertible readConverter, String filed, Class<?> filedClazz) {
        super();
        this.title = title;
        this.order = order;
        this.writeConverter = writeConverter;
        this.readConverter = readConverter;
        this.filed = filed;
        this.filedClazz = filedClazz;
    }
}
