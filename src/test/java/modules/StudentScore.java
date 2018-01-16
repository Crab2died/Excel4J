package modules;

import com.github.crab2died.annotation.ExcelField;

public class StudentScore {

    @ExcelField(title="学号", order = 1)
    private String num;

    @ExcelField(title="姓名", order = 2)
    private String name;

    @ExcelField(title="语文成绩", order = 3)
    private Double chinese ;

    @ExcelField(title="数学成绩", order = 4)
    private Double mathematics;

    @ExcelField(title="英语成绩", order = 5)
    private Double english;

    // 总成是excel内函数计算所得
    @ExcelField(title="总成绩", order = 1)
    private Double total;

    public String getNum() {
        return num;
    }

    public void setNum(String num) {
        this.num = num;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Double getChinese() {
        return chinese;
    }

    public void setChinese(Double chinese) {
        this.chinese = chinese;
    }

    public Double getMathematics() {
        return mathematics;
    }

    public void setMathematics(Double mathematics) {
        this.mathematics = mathematics;
    }

    public Double getEnglish() {
        return english;
    }

    public void setEnglish(Double english) {
        this.english = english;
    }

    public Double getTotal() {
        return total;
    }

    public void setTotal(Double total) {
        this.total = total;
    }

    @Override
    public String toString() {
        return "StudentScore{" +
                "num='" + num + '\'' +
                ", name='" + name + '\'' +
                ", chinese=" + chinese +
                ", mathematics=" + mathematics +
                ", english=" + english +
                ", total=" + total +
                '}';
    }
}
