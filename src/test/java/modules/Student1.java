package modules;


import com.github.crab2died.annotation.ExcelField;
import lombok.Data;

@Data
public class Student1 {

    // 学号
    @ExcelField(title = "学号", order = 1)
    private String id;

    // 姓名
    @ExcelField(title = "姓名", order = 2)
    private String name;

    // 班级
    @ExcelField(title = "班级", order = 3)
    private String classes;

    public Student1() {

    }

    public Student1(String id, String name, String classes) {
        this.id = id;
        this.name = name;
        this.classes = classes;
    }


}
