package base;


import com.github.ExcelUtils;
import moudles.Student;
import org.junit.Test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Module2Excel {

    @Test
    public void object2Excel() {
        String tempPath = "D:\\IdeaSpace\\Excel4J\\src\\test\\java\\resource\\template.xlsx";
        List<Student> list = new ArrayList<>();
        list.add(new Student("1010001", "盖伦", "六年级三班"));
        list.add(new Student("1010002", "古尔丹", "一年级三班"));
        list.add(new Student("1010003", "蒙多(被开除了)", "六年级一班"));
        list.add(new Student("1010004", "萝卜特", "三年级二班"));
        list.add(new Student("1010005", "奥拉基", "三年级二班"));
        list.add(new Student("1010006", "得嘞", "四年级二班"));
        list.add(new Student("1010007", "瓜娃子", "五年级一班"));
        Map<String, String> data = new HashMap<>();
        data.put("title", "战争学院花名册");
        data.put("info", "学校统一花名册");
        ExcelUtils.getInstance().exportObj2ExcelByTemplate(data, tempPath, "D:\\Q.xlsx", list, Student.class, false,
                true);
    }
}
