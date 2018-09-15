package converter;

import com.github.crab2died.constant.LanguageEnum;
import com.github.crab2died.converter.WriteConvertible;

/**
 * 导出excel Boolean 数据转换器
 */
public class Student2ExpelWriteConverter implements WriteConvertible {

    @Override
    public Object execWrite(Object object) {
        return execWrite(object, null);
    }

    @Override
    public Object execWrite(Object object, String language) {
        Boolean expel = (Boolean) object;
        if (!LanguageEnum.CHINESE.getValue().equalsIgnoreCase(language)) {
            return expel ? "Yes" : "No";
        } else {
            return expel ? "是" : "否";
        }
    }
}
