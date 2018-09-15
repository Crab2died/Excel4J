package com.github.crab2died.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface I18nField {

    /**
     * 国际化标题栏(例如: ["zh-cn|学生","en-us|student"])
     *
     * @return
     */
    String[] titles();

}
