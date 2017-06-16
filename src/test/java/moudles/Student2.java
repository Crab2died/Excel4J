package moudles;

import com.github.annotation.ExcelField;

import java.util.Date;

/**
 * <p></p></br>
 * author : wbhe2</br>
 * date  : 2017/6/15  15:15</br>
 */
public class Student2 {

    @ExcelField(title = "学号", order = 1)
    private Long id;

    @ExcelField(title = "姓名", order = 2)
    private String name;

    @ExcelField(title = "入学日期", order = 3)
    private Date date;

    @ExcelField(title = "班级", order = 4)
    private Integer classes;

    @ExcelField(title = "是否开除", order = 5)
    private String expel;

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

    public String getExpel() {
        return expel;
    }

    public void setExpel(String expel) {
        this.expel = expel;
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
