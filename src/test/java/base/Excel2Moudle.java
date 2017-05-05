package base;

import com.github.ExcelUtil;
import moudles.Student;
import org.junit.Test;

import java.util.List;

public class Excel2Moudle {

    @Test
    public void excel2Object(){
        String path = "D:\\IdeaSpace\\Excel4J\\src\\test\\java\\resource\\student.xlsx";

        System.out.println("读取全部：");
        List<Student> students = ExcelUtil.getInstance().readExcel2ObjsByClasspath(path, Student.class);
        for (Student stu : students){
            System.out.println(stu.getName() + " -- " + stu.getClasses());
        }

        System.out.println("读取指定行数：");
        students  = ExcelUtil.getInstance().readExcel2ObjsByClasspath(path, Student.class, 0, 2);
        for (Student stu : students){
            System.out.println(stu.getName() + " -- " + stu.getClasses());
        }
    }
}
