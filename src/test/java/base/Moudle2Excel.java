package base;


import com.github.ExcelUtil;
import moudles.Student;
import org.junit.Test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Moudle2Excel {

    @Test
    public void object2Excel(){
        String tempPath = "D:\\JProject\\Excel4J\\src\\test\\java\\resource\\template.xlsx";
        List<Student> list = new ArrayList<>();
        list.add(new Student("1010001", "盖伦", "六年级三班"));
        list.add(new Student("1010002", "古尔丹", "一年级三班"));
        list.add(new Student("1010003", "蒙多", "六年级一班"));
        list.add(new Student("1010004", "萝卜特", "三年级二班"));
        Map<String, String> datas = new HashMap<>();
        datas.put("title", "战争学院花名册");
        datas.put("info", "学校统一花名册");
        ExcelUtil.getInstance().exportObj2ExcelByTemplate(datas, tempPath, "D:\\Q.xlsx", list, Student.class, false, true);
    }
}
