package com.github.crab2died.constant;

import lombok.Getter;
import lombok.Setter;

/**
 * @author junmingyang
 */
public enum LanguageEnum {

    /**
     * 中文
     */
    CHINESE("中文", "zh-cn"),

    /**
     * 英语
     */
    ENGLISH("英语", "en-us");

    LanguageEnum(String name, String value) {
        this.name = name;
        this.value = value;
    }

    @Setter
    @Getter
    private String name;

    @Setter
    @Getter
    private String value;
}
