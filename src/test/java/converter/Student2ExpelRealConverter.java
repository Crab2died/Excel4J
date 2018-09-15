package converter;

import com.github.crab2died.constant.LanguageEnum;
import com.github.crab2died.converter.ReadConvertible;

/**
 * excel是否开除 列数据转换器
 */
public class Student2ExpelRealConverter implements ReadConvertible {

    @Override
    public Object execRead(String object) {
        return execRead(object, null);
    }

    @Override
    public Object execRead(String object, String language) {
        if (language.equalsIgnoreCase(LanguageEnum.CHINESE.getValue())) {
            return object.equals("是");
        }
        return object.equals("Yes");
    }
}
