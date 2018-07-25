package base;

import com.github.crab2died.ExcelUtils;
import modules.Student1;
import modules.Student2;
import modules.StudentScore;
import org.junit.Test;

import java.nio.file.Paths;
import java.util.List;

public class Excel2Module {

    @Test
    public void excel2Object() throws Exception {

        String path = "D:\\workspace\\Excel4J\\src\\test\\resources\\students_01.xlsx";

        System.out.println("读取全部：");
        List<Student1> students = ExcelUtils.getInstance().readExcel2Objects(path, Student1.class);
        for (Student1 stu : students) {
            System.out.println(stu);
        }
        System.out.println("读取指定行数：");
        students = ExcelUtils.getInstance().readExcel2Objects(path, Student1.class, 0, 3, 0);
        for (Student1 stu : students) {
            System.out.println(stu);
        }
    }

    @Test
    public void excel2Object2() {

        String path = "D:\\JProject\\Excel4J\\src\\test\\resources\\students_02.xlsx";
        try {

            // 1)
            // 不基于注解,将Excel内容读至List<List<String>>对象内
            List<List<String>> lists = ExcelUtils.getInstance().readExcel2List(path, 1, 2, 0);
            System.out.println("读取Excel至String数组：");
            for (List<String> list : lists) {
                System.out.println(list);
            }

            // 2)
            // 基于注解,将Excel内容读至List<Student2>对象内
            // 验证读取转换函数Student2ExpelConverter
            // 注解 `@ExcelField(title = "是否开除", order = 5, readConverter =  Student2ExpelConverter.class)`
            List<Student2> students = ExcelUtils.getInstance().readExcel2Objects(path, Student2.class, 0, 0);
            System.out.println("读取Excel至对象数组(支持类型转换)：");
            for (Student2 st : students) {
                System.out.println(st);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 基于注解,将Excel内容读至List<Student2>对象内
    // 验证读取转换函数Student2ExpelConverter，注解 `@ExcelField(title = "是否开除", order = 5, readConverter = Student2ExpelConverter.class)`
    @Test
    public void testReadConverter() throws Exception {

        String path = "D:\\workspace\\Excel4J\\src\\test\\resources\\students_02.xlsx";
        List<Student2> students = ExcelUtils.getInstance().readExcel2Objects(path, Student2.class, 0, 0);
        System.out.println("读取Excel至对象数组(支持类型转换)：");
        for (Student2 st : students) {
            System.out.println(st);
        }
    }

    //测试读取带有公式的单元格，并返回公式的值
    @Test
    public void testReadExcel_XLS() throws  Exception{
        String path = "D:\\workspace\\Excel4J\\src\\test\\resources\\StudentScore.xlsx";
        System.out.println(Paths.get(path).toUri().getPath());
        List<StudentScore> projectExcelModels = ExcelUtils.getInstance().readExcel2Objects(path, StudentScore.class);
        System.out.println(projectExcelModels);
    }
}
