package com.github.crab2died.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 国际化标题注解
 *
 * @author 菩提树下的杨过(http : / / yjmyzz.cnblogs.com)
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface I18nField {

    /**
     * 国际化标题栏(例如: ["zh-cn|学生","en-us|student"])
     *
     * @return 国际化标题配置数组
     */
    String[] titles();

}
