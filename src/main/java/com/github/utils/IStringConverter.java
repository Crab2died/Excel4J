package com.github.utils;

/**
 * 自定义字符串转换接口，用于将Excel导入的数据转换为自定义类型，例如将用';'分割的字符串转换为数组类型
 * @author XiaoYu
 *
 */
public interface IStringConverter {

	/**
	 * 实现此方法对传入的@value进行数据类型转换
	 * @param field 
	 * @param value
	 * @return
	 */
	Object convert(String field, String value);
}
