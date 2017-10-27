package converter;

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
        return DateUtils.date2Str(date, DateUtils.DATE_FORMAT_MSEC_T_Z);
    }
}
