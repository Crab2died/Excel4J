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
