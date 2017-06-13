package base;

import com.github.ExcelUtils;
import moudles.Student;
import org.junit.Test;

import java.util.List;

public class Excel2Module {

    @Test
    public void excel2Object() {
        String path = "D:\\IdeaSpace\\Excel4J\\src\\test\\java\\resource\\student.xlsx";

        System.out.println("读取全部：");
        List<Student> students = ExcelUtils.getInstance().readExcel2ObjByClasspath(path, Student.class);
        for (Student stu : students) {
            System.out.println(stu.getName() + " -- " + stu.getClasses());
        }

        System.out.println("读取指定行数：");
        students = ExcelUtils.getInstance().readExcel2ObjByClasspath(path, Student.class, 0, 2);
        for (Student stu : students) {
            System.out.println(stu.getName() + " -- " + stu.getClasses());
        }
    }
}
