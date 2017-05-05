package base;

import com.github.ExcelUtil;
import moudles.Student;
import org.junit.Test;

import java.util.List;

public class Excel2Moudle {

    @Test
    public void excel2Object(){
        String path = "D:\\IdeaSpace\\Excel4J\\src\\test\\java\\resource\\student.xlsx";
        List<Object> students = ExcelUtil.getInstance().readExcel2ObjsByClasspath(path, Student.class);
        for (Object obj : students){
            Student stu = (Student) obj;
            System.out.println(stu.getName() + " -- " + stu.getClasses());
        }

        students  = ExcelUtil.getInstance().readExcel2ObjsByClasspath(path, Student.class, 0, 2);
        for (Object obj : students){
            Student stu = (Student) obj;
            System.out.println(stu.getName() + " -- " + stu.getClasses());
        }
    }
}
