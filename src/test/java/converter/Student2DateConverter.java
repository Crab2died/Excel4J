package converter;

import com.github.crab2died.constant.LanguageEnum;
import com.github.crab2died.converter.WriteConvertible;
import com.github.crab2died.utils.DateUtils;

import java.util.Date;

/**
 * 导出excel日期数据转换器
 */
public class Student2DateConverter implements WriteConvertible {

    @Override
    public Object execWrite(Object object) {
        Date date = (Date) object;
        if (language.equalsIgnoreCase(LanguageEnum.CHINESE.getValue())) {
            return DateUtils.date2Str(date, DateUtils.DATE_FORMAT_DAY);
        } else {
            return DateUtils.date2Str(date, DateUtils.DATE_FORMAT_DAY_2);
        }
    }
}
