package com.github.handler;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

/**
 * <p></p></br>
 * author : wbhe2</br>
 * date  : 2017/6/13  11:24</br>
 */
public class ExcelTemplate {

    /**
     * excel工作簿
     */
    private Workbook workbook;
    /**
     * excel工作表
     */
    private Sheet sheet;
    /**
     * 当前行对象
     */
    private Row curRow;
    /**
     * 默认样式
     */
    private CellStyle defaultStyle;
    /**
     * 指定行样式
     */
    private Map<Integer, CellStyle> appointLineStyle;
    /**
     * 单数行样式
     */
    private CellStyle singleLineStyle;
    /**
     * 双数行样式
     */
    private CellStyle doubleLineStyle;
    /**
     * 工作表sheet号
     */
    private int sheetIndex;
    /**
     * 数据的初始化列数
     */
    private int initColumnIndex;
    /**
     * 数据的初始化行数
     */
    private int initRowIndex;
    /**
     * 当前列数
     */
    private int curColumnIndex;
    /**
     * 当前行数
     */
    private int curRowIndex;
    /**
     * 最后一行的数据
     */
    private int lastRowIndex;
    /**
     * 默认行高
     */
    private float rowHeight;
    /**
     * 序号
     */
    private int serialNumber;


    private ExcelTemplate() {
    }

    private ExcelTemplate(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    static public ExcelTemplate instance() {
        return new ExcelTemplate();
    }

    static public ExcelTemplate instance(int sheetIndex) {
        return new ExcelTemplate(sheetIndex);
    }

    /**
     * 从某个路径来读取模板
     *
     * @param path 模板路径
     * @return ExcelTemplate
     */
    public ExcelTemplate readTemplateByPath(String path) {
        try {
            this.workbook = WorkbookFactory.create(new File(path));
            initTemplate();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
            throw new RuntimeException("读取模板格式有错！请检查");
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("读取模板不存在！请检查");
        }
        return this;
    }

    /**
     * 从classpath路径下读取相应的模板文件
     *
     * @param path 模板路径
     * @return ExcelTemplate
     */
    public ExcelTemplate readTemplateByClasspath(String path) {
        try {
            this.workbook = WorkbookFactory.create(ExcelTemplate.class.getResourceAsStream(path));
            initTemplate();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
            throw new RuntimeException("读取模板格式有错！请检查");
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("读取模板不存在！请检查");
        }
        return this;
    }

    /**
     * 将文件写到相应的路径下
     *
     * @param filepath 输出文件路径
     */
    public void writeToFile(String filepath) {
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(filepath);
            this.workbook.write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw new RuntimeException("写入的文件不存在");
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("写入数据失败:" + e.getMessage());
        } finally {
            try {
                if (fos != null)
                    fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 将文件写到某个输出流中
     *
     * @param os 输出流
     */
    public void writeToStream(OutputStream os) {
        try {
            this.workbook.write(os);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("写入流失败:" + e.getMessage());
        }
    }

    private void initTemplate() {
        this.sheet = this.workbook.getSheetAt(this.sheetIndex);
        initConfigData();
        this.lastRowIndex = this.sheet.getLastRowNum();
        this.curRow = this.sheet.createRow(this.curRowIndex);
    }


    /**
     * 初始化数据信息
     */
    private void initConfigData() {
        boolean findData = false;
        boolean findSer = false;
        for (Row row : sheet) {
            if (findData)
                break;
            for (Cell c : row) {
                if (c.getCellType() != Cell.CELL_TYPE_STRING)
                    continue;
                String str = c.getStringCellValue().trim();
                if (str.equals(HanderConstant.SERIAL_NUMBER)) {
                    this.serialNumber = c.getColumnIndex();
                    findSer = true;
                }
                if (str.equals(HanderConstant.DATA_INIT_INDEX)) {
                    this.initColumnIndex = c.getColumnIndex();
                    this.initRowIndex = row.getRowNum();
                    this.curColumnIndex = this.initColumnIndex;
                    this.curRowIndex = this.initRowIndex;
                    findData = true;
                    this.defaultStyle = c.getCellStyle();
                    this.rowHeight = row.getHeightInPoints();
                    initStyles();
                    break;
                }
            }
        }
        if (!findSer) {
            initSer();
        }
    }

    /**
     * 初始化样式信息
     */
    private void initStyles() {
        this.appointLineStyle = new HashMap<>();
        for (Row row : sheet) {
            for (Cell c : row) {
                if (c.getCellType() != Cell.CELL_TYPE_STRING)
                    continue;
                String str = c.getStringCellValue();
                if (null != str && !"".equals(str))
                    str = str.trim().toLowerCase();
                if (HanderConstant.DEFAULT_STYLE.equals(str)) {
                    this.defaultStyle = c.getCellStyle();
                    clearCell(c);
                }
                if (HanderConstant.APPOINT_LINE_STYLE.equals(str)) {
                    this.appointLineStyle.put(c.getRowIndex(), c.getCellStyle());
                    clearCell(c);
                }
                if (HanderConstant.SINGLE_LINE_STYLE.equals(str)) {
                    this.singleLineStyle = c.getCellStyle();
                    clearCell(c);
                }
                if (HanderConstant.DOUBLE_LINE_STYLE.equals(str)) {
                    this.doubleLineStyle = c.getCellStyle();
                    clearCell(c);
                }
            }
        }
    }

    private void clearCell(Cell cell){
        cell.setCellStyle(null);
        cell.setCellValue("");
    }

    /**
     * 初始化序号位置
     */
    private void initSer() {
        for (Row row : sheet) {
            for (Cell c : row) {
                if (c.getCellType() != Cell.CELL_TYPE_STRING)
                    continue;
                String str = c.getStringCellValue().trim();
                if (HanderConstant.SERIAL_NUMBER.equals(str)) {
                    this.serialNumber = c.getColumnIndex();
                }
            }
        }
    }
    /**
     * 插入序号，会自动找相应的序号标示的位置完成插入
     */
    public void insertSer() {
        int index = 1;
        Row row;
        Cell c;
        for (int i = this.initRowIndex; i < this.curRowIndex; i++) {
            row = this.sheet.getRow(i);
            c = row.createCell(this.serialNumber);
            setCellStyle(c);
            c.setCellValue(index++);
        }
    }

    /**
     * 创建新行，在使用时只要添加完一行，需要调用该方法创建
     */
    public void createNewRow() {
        if (this.lastRowIndex > this.curRowIndex && this.curRowIndex != this.initRowIndex) {
            this.sheet.shiftRows(this.curRowIndex, this.lastRowIndex, 1, true, true);
            this.lastRowIndex++;
        }
        this.curRow = this.sheet.createRow(this.curRowIndex);
        this.curRow.setHeightInPoints(this.rowHeight);
        this.curRowIndex++;
        this.curColumnIndex = this.initColumnIndex;
    }

    /**
     * 根据map替换相应的常量，通过Map中的值来替换#开头的值
     *
     * @param data 替换映射
     */
    public void replaceFinalData(Map<String, String> data) {
        if (data == null)
            return;
        for (Row row : this.sheet) {
            for (Cell c : row) {
                if (c.getCellType() != Cell.CELL_TYPE_STRING)
                    continue;
                String str = c.getStringCellValue().trim();
                if (str.startsWith("#")) {
                    if (data.containsKey(str.substring(1))) {
                        c.setCellValue(data.get(str.substring(1)));
                    }
                }
            }
        }
    }
    /**
     * 基于Properties的替换，依然也是替换#开始的
     *
     * @param prop Properties映射
     */
    public void replaceFinalData(Properties prop) {
        if (prop == null)
            return;
        for (Row row : this.sheet) {
            for (Cell c : row) {
                if (c.getCellType() != Cell.CELL_TYPE_STRING)
                    continue;
                String str = c.getStringCellValue().trim();
                if (str.startsWith("#")) {
                    if (prop.containsKey(str.substring(1))) {
                        c.setCellValue(prop.getProperty(str.substring(1)));
                    }
                }
            }
        }
    }
    /**
     * 设置某个元素的样式
     *
     * @param cell cell元素
     */
    private void setCellStyle(Cell cell) {
        if (null != this.appointLineStyle && this.appointLineStyle.containsKey(cell.getRowIndex())) {
            cell.setCellStyle(this.appointLineStyle.get(cell.getRowIndex()));
            return;
        }
        if (null != this.singleLineStyle && (cell.getRowIndex() % 2 != 0)) {
            cell.setCellStyle(this.singleLineStyle);
            return;
        }
        if (null != this.doubleLineStyle && (cell.getRowIndex() % 2 == 0)) {
            cell.setCellStyle(this.doubleLineStyle);
            return;
        }
        if (null != this.defaultStyle)
            cell.setCellStyle(this.defaultStyle);
    }


    /**
     * <p>设置Excel元素样式及内容</p></br>
     */
    public void createCell(Object value) {
        Cell cell = this.curRow.createCell(curColumnIndex);
        setCellStyle(cell);
        if (null == value || "".equals(value)) {
            this.curColumnIndex++;
            return;
        }

        if (String.class == value.getClass()) {
            cell.setCellValue((String) value);
            this.curColumnIndex++;
            return;
        }

        if (int.class == value.getClass()) {
            cell.setCellValue((int) value);
            this.curColumnIndex++;
            return;
        }

        if (Integer.class == value.getClass()) {
            cell.setCellValue((Integer) value);
            this.curColumnIndex++;
            return;
        }

        if (double.class == value.getClass()) {
            cell.setCellValue((double) value);
            this.curColumnIndex++;
            return;
        }

        if (Double.class == value.getClass()) {
            cell.setCellValue((Double) value);
            this.curColumnIndex++;
            return;
        }

        if (Date.class == value.getClass()) {
            cell.setCellValue((Date) value);
            this.curColumnIndex++;
            return;
        }

        if (boolean.class == value.getClass()) {
            cell.setCellValue((boolean) value);
            this.curColumnIndex++;
            return;
        }
        if (Boolean.class == value.getClass()) {
            cell.setCellValue((Boolean) value);
            this.curColumnIndex++;
            return;
        }
        if (Calendar.class == value.getClass()) {
            cell.setCellValue((Calendar) value);
            this.curColumnIndex++;
            return;
        }
        this.curColumnIndex++;
    }
}
