package modules;

import com.github.crab2died.annotation.ExcelField;
import converter.Student2DateConverter;
import converter.Student2ExpelConverter;

import java.util.Date;


public class Student2 {

    @ExcelField(title = "学号", order = 1)
    private Long id;

    @ExcelField(title = "姓名", order = 2)
    private String name;

    // 写入数据转换器 Student2DateConverter
    @ExcelField(title = "入学日期", order = 3, writeConverter = Student2DateConverter.class)
    private Date date;

    @ExcelField(title = "班级", order = 4)
    private Integer classes;

    // 读取数据转换器 Student2ExpelConverter
    @ExcelField(title = "是否开除", order = 5, readConverter = Student2ExpelConverter.class)
    private boolean expel;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public Integer getClasses() {
        return classes;
    }

    public void setClasses(Integer classes) {
        this.classes = classes;
    }

    public boolean isExpel() {
        return expel;
    }

    public void setExpel(boolean expel) {
        this.expel = expel;
    }

    public Student2(Long id, String name, Date date, Integer classes, boolean expel) {
        this.id = id;
        this.name = name;
        this.date = date;
        this.classes = classes;
        this.expel = expel;
    }

    public Student2() {
    }

    @Override
    public String toString() {
        return "Student2{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", date=" + date +
                ", classes=" + classes +
                ", expel='" + expel + '\'' +
                '}';
    }
}
