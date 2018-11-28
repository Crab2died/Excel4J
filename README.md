```
                            ___________                   .__      _____      ____.
                            \_   _____/__  ___ ____  ____ |  |    /  |  |    |    |
                             |    __)_\  \/  // ___\/ __ \|  |   /   |  |_   |    |
                             |        \>    <\  \__\  ___/|  |__/    ^   /\__|    |
                            /_______  /__/\_ \\___  >___  >____/\____   |\________|
                                    \/      \/    \/    \/           |__|          
                                                             (version: 3.0.0-Alpha)
```
---

![version](https://img.shields.io/badge/version-3.0.0--Alpha-green.svg) 
[![GitHub license](https://img.shields.io/github/license/Crab2died/Excel4J.svg)](https://github.com/Crab2died/Excel4J/blob/master/LICENSE)
[![Maven Central](https://img.shields.io/maven-central/v/org.apache.maven/apache-maven.svg)](https://search.maven.org/search?q=a:Excel4J)

> 紧急修复以绝对路径指定模板来导出会导致模板被修改的BUG,以及读取Excel数据会修改原Excel文件,建议升级至2.1.4-Final2版本

## 一. 更新记录
### 1. v3.x
   1. 新增CSV(包含基于ExcelField注解)的导出支持
   2. 新增CSV(包含基于ExcelField注解)的导入支持

### 2. v2.x
   1. Excel读取支持部分类型转换了(如转为Integer,Long,Date(部分)等) v2.0.0之前只能全部内容转为String
   2. Excel支持非注解读取Excel内容了,内容存于`List<List<String>>`对象内
   3. 现在支持`List<List<String>>`导出Excel了(可以不基于模板)
   4. Excel新增了Map数据样式映射功能(模板可为每个key设置一个样式,定义为:&key, 导出Map数据的样式将与key值映射)
   5. 新增读取Excel数据转换器接口`com.github.converter.ReadConvertible`
   6. 新增写入Excel数据转换器接口`com.github.converter.WriteConvertible`
   7. 支持多sheet一键导出，多sheet导出封装Wrapper详见`com.github.sheet.wrapper`包内包装类
   8. 修复以绝对路径指定模板来导出会导致模板被修改的BUG,以及读取Excel数据会修改原Excel文件,建议升级至2.1.4-Final2版本
   9. 修复已知bug及代码与注释优化

## 二. 基于注解(/src/test/java/modules/Student2.java)
```
    @ExcelField(title = "学号", order = 1)
    private Long id;

    @ExcelField(title = "姓名", order = 2)
    private String name;

    // 写入数据转换器 Student2DateConverter
    @ExcelField(title = "入学日期", order = 3, writeConverter = Student2DateConverter.class)
    private Date date;

    @ExcelField(title = "班级", order = 4)
    private Integer classes;

    // 读取数据转换器 Student2ExpelConverter
    @ExcelField(title = "是否开除", order = 5, readConverter = Student2ExpelConverter.class)
    private boolean expel;
```

## 三. 读取Excel快速实现

### 1.待读取Excel(截图)
![待读取Excel截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.0.0/students_02.png)

### 2. 读取转换器(/src/test/java/converter/Student2ExpelConverter.java)
```
    /**
     * excel是否开除 列数据转换器
     */
    public class Student2ExpelConverter implements ReadConvertible{
    
        @Override
        public Object execRead(String object) {
    
            return object.equals("是");
        }
    }
```

### 3. 读取函数(/src/test/java/base/Excel2Module.java#excel2Object2)
```
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
```

### 4. 读取结果
```
    读取Excel至String数组：
    [10000000000001, 张三, 2016/01/19, 101, 是]
    [10000000000002, 李四, 2017-11-17 10:19:10, 201, 否]
    读取Excel至对象数组(支持类型转换)：
    Student2{id=10000000000001, name='张三', date=Tue Jan 19 00:00:00 CST 2016, classes=101, expel='true'}
    Student2{id=10000000000002, name='李四', date=Fri Nov 17 10:19:10 CST 2017, classes=201, expel='false'}
    Student2{id=10000000000004, name='王二', date=Fri Nov 17 00:00:00 CST 2017, classes=301, expel='false'}
```

## 四. 导出Excel

### 1. 不基于模板快速导出

#### 1) 导出函数(/src/test/java/base/Module2Excel.java#testList2Excel)
```
    @Test
    public void testList2Excel() throws Exception {
        
        List<List<String>> list2 = new ArrayList<>();
        List<String> header = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            List<String> _list = new ArrayList<>();
            for (int j = 0; j < 10; j++) {
                _list.add(i + " -- " + j);
            }
            list2.add(_list);
            header.add(i + "---");
        }
        ExcelUtils.getInstance().exportObjects2Excel(list2, header, "D:/D.xlsx");
    }
```
#### 2) 导出效果(截图)
![无模板导出截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.0.0/list_export.png)

### 2. 带有写入转换器函数的导出

#### 1) 转换器(/src/test/java/converter/Student2DateConverter.java)
```
    /**
     * 导出excel日期数据转换器
     */
    public class Student2DateConverter implements WriteConvertible {
    
    
        @Override
        public Object execWrite(Object object) {
    
            Date date = (Date) object;
            return DateUtils.date2Str(date, DateUtils.DATE_FORMAT_MSEC_T_Z);
        }
    }

```
#### 2）导出函数(/src/test/java/base/Module2Excel.java#testWriteConverter)
```
    // 验证日期转换函数 Student2DateConverter
    // 注解 `@ExcelField(title = "入学日期", order = 3, writeConverter = Student2DateConverter.class)`
    @Test
    public void testWriteConverter() throws Exception {

        List<Student2> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            list.add(new Student2(10000L + i, "学生" + i, new Date(), 201, false));
        }
        ExcelUtils.getInstance().exportObjects2Excel(list, Student2.class, true, "sheet0", true, "D:/D.xlsx");
    }
```
#### 3) 导出效果(截图)
![无模板导出截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.0.0/converter_export.png)

### 3. 基于模板`List<Oject>`导出

#### 1) 导出函数(/src/test/java/base/Module2Excel.java#testObject2Excel)
```
    @Test
    public void testObject2Excel() throws Exception {

        String tempPath = "/normal_template.xlsx";
        List<Student1> list = new ArrayList<>();
        list.add(new Student1("1010001", "盖伦", "六年级三班"));
        list.add(new Student1("1010002", "古尔丹", "一年级三班"));
        list.add(new Student1("1010003", "蒙多(被开除了)", "六年级一班"));
        list.add(new Student1("1010004", "萝卜特", "三年级二班"));
        list.add(new Student1("1010005", "奥拉基", "三年级二班"));
        list.add(new Student1("1010006", "得嘞", "四年级二班"));
        list.add(new Student1("1010007", "瓜娃子", "五年级一班"));
        list.add(new Student1("1010008", "战三", "二年级一班"));
        list.add(new Student1("1010009", "李四", "一年级一班"));
        Map<String, String> data = new HashMap<>();
        data.put("title", "战争学院花名册");
        data.put("info", "学校统一花名册");
        // 基于模板导出Excel
        ExcelUtils.getInstance().exportObjects2Excel(tempPath, 0, list, data, Student1.class, false, "D:/A.xlsx");
        // 不基于模板导出Excel
        ExcelUtils.getInstance().exportObjects2Excel(list, Student1.class, true, null, true, "D:/B.xlsx");

    }
```

#### 2) 导出模板(截图)
![导出模板截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.0.0/normal_template.png)

#### 3) 基于模板导出结果(截图)
![基于模板导出结果图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.0.0/normal_export.png)

#### 4) 不基于模板导出结果(截图)
![不基于模板导出结果图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.0.0/object_export.png)

### 4. 基于模板`Map<String, Collection<Object.toString>>`导出

#### 1) 导出函数(/src/test/java/base/Module2Excel.java#testMap2Excel)
```
    @Test
    public void testMap2Excel() throws Exception {

        Map<String, List> classes = new HashMap<>();

        Map<String, String> data = new HashMap<>();
        data.put("title", "战争学院花名册");
        data.put("info", "学校统一花名册");

        classes.put("class_one", new ArrayList<Student1>() {{
            add(new Student1("1010009", "李四", "一年级一班"));
            add(new Student1("1010002", "古尔丹", "一年级三班"));
        }});
        classes.put("class_two", new ArrayList<Student1>() {{
            add(new Student1("1010008", "战三", "二年级一班"));
        }});
        classes.put("class_three", new ArrayList<Student1>() {{
            add(new Student1("1010004", "萝卜特", "三年级二班"));
            add(new Student1("1010005", "奥拉基", "三年级二班"));
        }});
        classes.put("class_four", new ArrayList<Student1>() {{
            add(new Student1("1010006", "得嘞", "四年级二班"));
        }});
        classes.put("class_six", new ArrayList<Student1>() {{
            add(new Student1("1010001", "盖伦", "六年级三班"));
            add(new Student1("1010003", "蒙多", "六年级一班"));
        }});

        ExcelUtils.getInstance().exportObject2Excel("/map_template.xlsx",
                0, classes, data, Student1.class, false, "D:/C.xlsx");
    }
```

#### 2) 导出模板(截图)
![导出模板截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.0.0/map_export_template.png)

#### 3) 导出结果(截图)
![导出结果图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.0.0/map_export.png)

## 五. Excel模板自定义属性,不区分大小写
### 1)  具体代码定义详见(/src/main/java/com/github/crab2died/handler/HandlerConstant)
### 2)  Excel模板自定义属性,不区分大小写
|       定义符        |      描述      |优先级(大到小)|
|:-------------------|:---------------|:----------:|
|$appoint_line_style |当前行样式       |       3    |
|$single_line_style  |单行样式         |       2    |
|$double_line_style  |双行样式         |       2    |
|$default_style      |默认样式         |       1    |
|$data_index         |数据插入的起始位置|       -    |
|$serial_number      |插入序号标记     |       -    |
    
## 六. 多sheet数据导出
### 1. 多sheet数据导出包装类,详见`com.github.sheet.wrapper`包内包装类
   多sheet数据导出只需将待导出数据封装入`com.github.sheet.wrapper`包内的Wrapper类即可实现多sheet一键导出

### 2. 无模板、无注解的多sheet导出`com.github.sheet.wrapper.SimpleSheetWrapper`
#### 1) 调用方法
```
    // 多sheet无模板、无注解导出
    @Test
    public void testBatchSimple2Excel() throws Exception {

        // 生成sheet数据
        List<SimpleSheetWrapper> list = new ArrayList<>();
        for (int i = 0; i <= 2; i++) {
            //表格内容数据
            List<String[]> data = new ArrayList<>();
            for (int j = 0; j < 1000; j++) {

                // 行数据(此处是数组) 也可以是List数据
                String[] rows = new String[5];
                for (int r = 0; r < 5; r++) {
                    rows[r] = "sheet_" + i + "row_" + j + "column_" + r;
                }
                data.add(rows);
            }
            // 表头数据
            List<String> header = new ArrayList<>();
            for (int h = 0; h < 5; h++) {
                header.add("column_" + h);
            }
            list.add(new SimpleSheetWrapper(data, header, "sheet_" + i));
        }
        ExcelUtils.getInstance().simpleSheet2Excel(list, "K.xlsx");
    }
```
#### 2) 导出结果(截图)
![导出结果截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.1.4/simple_wrapper.png)

### 3. 无模板、基于注解的多sheet导出`com.github.sheet.wrapper.NoTemplateSheetWrapper`
#### 1) 调用方法
``` 
    // 多sheet无模板、基于注解的导出
    @Test
    public void testBatchNoTemplate2Excel() throws Exception {

        List<NoTemplateSheetWrapper> sheets = new ArrayList<>();

        for (int s = 0; s < 3; s++) {
            List<Student2> list = new ArrayList<>();
            for (int i = 0; i < 1000; i++) {
                list.add(new Student2(10000L + i, "学生" + i, new Date(), 201, false));
            }
            sheets.add(new NoTemplateSheetWrapper(list, Student2.class, true, "sheet_" + s));
        }
        ExcelUtils.getInstance().noTemplateSheet2Excel(sheets, "EE.xlsx");
    }
```
### 2) 导出结果(截图)
![导出结果截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.1.4/no_template_wrapper.png)

### 4. 基于模板、注解的多sheet导出`com.github.sheet.wrapper.NormalSheetWrapper`
#### 1) 调用方法(注:为了测试方便，各个sheet数据相同)
``` 
    // 基于模板、注解的多sheet导出
    @Test
    public void testObject2BatchSheet() throws Exception {

        List<NormalSheetWrapper> sheets = new ArrayList<>();
        for (int i = 0; i < 2; i++) {
            List<Student1> list = new ArrayList<>();
            list.add(new Student1("1010001", "盖伦", "六年级三班"));
            list.add(new Student1("1010002", "古尔丹", "一年级三班"));
            list.add(new Student1("1010003", "蒙多(被开除了)", "六年级一班"));
            list.add(new Student1("1010004", "萝卜特", "三年级二班"));
            list.add(new Student1("1010005", "奥拉基", "三年级二班"));
            list.add(new Student1("1010006", "得嘞", "四年级二班"));
            list.add(new Student1("1010007", "瓜娃子", "五年级一班"));
            list.add(new Student1("1010008", "战三", "二年级一班"));
            list.add(new Student1("1010009", "李四", "一年级一班"));
            Map<String, String> data = new HashMap<>();
            data.put("title", "战争学院花名册");
            data.put("info", "学校统一花名册");
            sheets.add(new NormalSheetWrapper(i, list, data, Student1.class, false));
        }

        String tempPath = "/normal_batch_sheet_template.xlsx";

        // 基于模板导出Excel
        ExcelUtils.getInstance().normalSheet2Excel(sheets, tempPath, "AA.xlsx");
    }
```
#### 2) 导出模板(截图) (注:为了测试方便，模板样式大致相同，单元格颜色有区别)
   1. sheet1模板  
   ![sheet1模板截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.1.4/normal_template_sheet1.png)
   2. sheet2模板  
   ![sheet2模板截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.1.4/normal_template_sheet2.png)
#### 3) 导出结果(截图)
   1. sheet1导出结果  
   ![sheet1导出结果截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.1.4/normal_wrapper_sheet1.png)
   2. sheet2导出结果  
   ![sheet2导出结果截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.1.4/normal_wrapper_sheet2.png)

### 5. 形如`Map<String, Collection<Object.toString>>`数据基于模板、注解的多sheet导出`com.github.sheet.wrapper.MapSheetWrapper`
#### 1) 调用方法(注:为了测试方便，各个sheet数据相同)
``` 
    // Map数据的多sheet导出
    @Test
    public void testMap2BatchSheet() throws Exception {

        List<MapSheetWrapper> sheets = new ArrayList<>();

        for (int i = 0; i < 2; i++) {
            Map<String, List<?>> classes = new HashMap<>();

            Map<String, String> data = new HashMap<>();
            data.put("title", "战争学院花名册");
            data.put("info", "学校统一花名册");

            classes.put("class_one", Arrays.asList(
                    new Student1("1010009", "李四", "一年级一班"),
                    new Student1("1010002", "古尔丹", "一年级三班")
            ));
            classes.put("class_two", Collections.singletonList(
                    new Student1("1010008", "战三", "二年级一班")
            ));
            classes.put("class_three", Arrays.asList(
                    new Student1("1010004", "萝卜特", "三年级二班"),
                    new Student1("1010005", "奥拉基", "三年级二班")
            ));
            classes.put("class_four", Collections.singletonList(
                    new Student1("1010006", "得嘞", "四年级二班")
            ));
            classes.put("class_six", Arrays.asList(
                    new Student1("1010001", "盖伦", "六年级三班"),
                    new Student1("1010003", "蒙多", "六年级一班")
            ));

            sheets.add(new MapSheetWrapper(i, classes, data, Student1.class, false));
        }
        ExcelUtils.getInstance().mapSheet2Excel(sheets, "/map_batch_sheet_template.xlsx", "CC.xlsx");
    }
```
### 2) 导出模板(截图) (注:为了测试方便，模板样式大致相同，单元格颜色有区别)
   1. sheet1模板  
   ![sheet1模板截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.1.4/map_template_sheet1.png)
   2. sheet2模板  
   ![sheet2模板截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.1.4/map_template_sheet2.png)
#### 3) 导出结果(截图)
   1. sheet1导出结果  
   ![sheet1导出结果截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.1.4/map_wrapper_sheet1.png)
   2. sheet2导出结果  
   ![sheet2导出结果截图](https://raw.githubusercontent.com/Crab2died/Excel4J/master/src/test/resources/image/v2.1.4/map_wrapper_sheet2.png)

## 七. CSV文件的操作(完全支持ExcelField注解的所有配置)
### 1. 基于注解读取CSV文件
#### 1) 调用方法
```
    // 测试读取CSV文件
    @Test
    public void testReadCSV() throws Excel4JException {
        List<Student2> list = ExcelUtils.getInstance().readCSV2Objects("J.csv", Student2.class);
        System.out.println(list);
    } 
```
#### 2) 读取结果
```
    Student2{id=1000001, name='张三', date=Wed Nov 28 15:11:12 CST 2018, classes=1, expel='false'}
    Student2{id=1010002, name='古尔丹', date=Wed Nov 28 15:11:12 CST 2018, classes=2, expel='false'}
    Student2{id=1010003, name='蒙多(被开除了)', date=Wed Nov 28 15:11:12 CST 2018, classes=6, expel='false'}
    Student2{id=1010004, name='萝卜特', date=Wed Nov 28 15:11:12 CST 2018, classes=3, expel='false'}
    Student2{id=1010005, name='奥拉基', date=Wed Nov 28 15:11:12 CST 2018, classes=4, expel='false'}
    Student2{id=1010006, name='得嘞', date=Wed Nov 28 15:11:12 CST 2018, classes=4, expel='false'}
    Student2{id=1010007, name='瓜娃子', date=Wed Nov 28 15:11:12 CST 2018, classes=5, expel='false'}
    Student2{id=1010008, name='战三', date=Wed Nov 28 15:11:12 CST 2018, classes=4, expel='false'}
    Student2{id=1010009, name='李四', date=Wed Nov 28 15:11:12 CST 2018, classes=2, expel='false'}
```

### 2. 基于注解导出CSV文件
#### 1) 调用方法
```
    // 导出csv
    @Test
    public void testExport2CSV() throws Excel4JException {

        List<Student2> list = new ArrayList<>();
        list.add(new Student2(1000001L, "张三", new Date(), 1, true));
        list.add(new Student2(1010002L, "古尔丹", new Date(), 2, false));
        list.add(new Student2(1010003L, "蒙多(被开除了)", new Date(), 6, true));
        list.add(new Student2(1010004L, "萝卜特", new Date(), 3, false));
        list.add(new Student2(1010005L, "奥拉基", new Date(), 4, false));
        list.add(new Student2(1010006L, "得嘞", new Date(), 4, false));
        list.add(new Student2(1010007L, "瓜娃子", new Date(), 5, true));
        list.add(new Student2(1010008L, "战三", new Date(), 4, false));
        list.add(new Student2(1010009L, "李四", new Date(), 2, false));

        ExcelUtils.getInstance().exportObjects2CSV(list, Student2.class, "J.csv");
    }

    // 超大数据量导出csv
    // 9999999数据本地测试小于1min
    @Test
    public void testExport2CSV2() throws Excel4JException {

        List<Student2> list = new ArrayList<>();
        for (int i = 0; i < 9999999; i++) {
            list.add(new Student2(1000001L + i, "路人 -" + i, new Date(), i % 6, true));
        }
        ExcelUtils.getInstance().exportObjects2CSV(list, Student2.class, "L.csv");
    }
```
#### 2) 导出结果
```
    // 以下为导出CSV文件内容
    
    学号,姓名,入学日期,班级,是否开除
    1000001,张三,2018-11-28T15:11:12.815Z,1,true
    1010002,古尔丹,2018-11-28T15:11:12.815Z,2,false
    1010003,蒙多(被开除了),2018-11-28T15:11:12.815Z,6,true
    1010004,萝卜特,2018-11-28T15:11:12.815Z,3,false
    1010005,奥拉基,2018-11-28T15:11:12.815Z,4,false
    1010006,得嘞,2018-11-28T15:11:12.815Z,4,false
    1010007,瓜娃子,2018-11-28T15:11:12.815Z,5,true
    1010008,战三,2018-11-28T15:11:12.815Z,4,false
    1010009,李四,2018-11-28T15:11:12.815Z,2,false
```


## 八. 使用(JDK1.7及以上)
#### 1) github拷贝项目
```bash
>> git clone https://github.com/Crab2died/Excel4J.git Excel4J
>> package.cmd
```

#### 2) 最新版本maven引用：
```xml
<dependency>
    <groupId>com.github.crab2died</groupId>
    <artifactId>Excel4J</artifactId>
    <version>2.1.4-Final2</version>
</dependency>
```

## 九. 链接
#### github -> [github地址](https://github.com/Crab2died/Excel4J)
#### 码云(gitee) -> [码云地址](https://gitee.com/Crab2Died/Excel4J)