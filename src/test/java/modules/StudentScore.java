package modules;

import com.github.crab2died.annotation.ExcelField;
import lombok.Data;

@Data
public class StudentScore {

    @ExcelField(title = "学号", order = 1)
    private String num;

    @ExcelField(title = "姓名", order = 2)
    private String name;

    @ExcelField(title = "语文成绩", order = 3)
    private Double chinese;

    @ExcelField(title = "数学成绩", order = 4)
    private Double mathematics;

    @ExcelField(title = "英语成绩", order = 5)
    private Double english;

    // 总成是excel内函数计算所得
    @ExcelField(title = "总成绩", order = 1)
    private Double total;


}
