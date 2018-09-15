package modules;

import com.github.crab2died.annotation.ExcelField;
import com.github.crab2died.annotation.I18nField;
import converter.Student2DateConverter;
import converter.Student2ExpelRealConverter;
import converter.Student2ExpelWriteConverter;
import lombok.Data;

import java.util.Date;

@Data
public class Student2 {

    @ExcelField(title = "学号", order = 1)
    @I18nField(titles = {"en-us|student id"})
    private Long id;

    @ExcelField(title = "姓名", order = 2)
    @I18nField(titles = {"en-us|name"})
    private String name;

    // 写入数据转换器 Student2DateConverter
    @ExcelField(title = "入学日期", order = 3, writeConverter = Student2DateConverter.class)
    @I18nField(titles = {"en-us|enroll date"})
    private Date date;

    @ExcelField(title = "班级", order = 4)
    @I18nField(titles = {"en-us|class"})
    private Integer classes;

    // 读写数据转换器 Student2ExpelRealConverter
    @ExcelField(title = "是否开除", order = 5, readConverter = Student2ExpelRealConverter.class, writeConverter = Student2ExpelWriteConverter.class)
    @I18nField(titles = {"en-us|is expel?"})
    private boolean expel;

    public Student2() {

    }

    public Student2(Long id, String name, Date date, Integer classes, boolean expel) {
        this.id = id;
        this.name = name;
        this.date = date;
        this.classes = classes;
        this.expel = expel;
    }


}
