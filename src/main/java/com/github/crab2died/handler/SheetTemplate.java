package com.github.crab2died.handler;

import com.github.crab2died.exceptions.Excel4JException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

public class SheetTemplate implements Closeable {

    /**
     * 当前工作簿
     */
    Workbook workbook;
    /**
     * 当前工作sheet表
     */
    Sheet sheet;
    /**
     * 当前表编号
     */
    int sheetIndex;
    /**
     * 当前行
     */
    Row currentRow;
    /**
     * 当前列数
     */
    int currentColumnIndex;
    /**
     * 当前行数
     */
    int currentRowIndex;
    /**
     * 默认样式
     */
    CellStyle defaultStyle;
    /**
     * 指定行样式
     */
    Map<Integer, CellStyle> appointLineStyle = new HashMap<>();
    /**
     * 分类样式模板
     */
    Map<String, CellStyle> classifyStyle = new HashMap<>();
    /**
     * 单数行样式
     */
    CellStyle singleLineStyle;
    /**
     * 双数行样式
     */
    CellStyle doubleLineStyle;
    /**
     * 数据的初始化列数
     */
    int initColumnIndex;
    /**
     * 数据的初始化行数
     */
    int initRowIndex;

    /**
     * 最后一行的数据
     */
    int lastRowIndex;
    /**
     * 默认行高
     */
    float rowHeight;
    /**
     * 序号坐标点
     */
    int serialNumberColumnIndex = -1;
    /**
     * 当前序号
     */
    int serialNumber;

    /*-----------------------------------写出数据开始-----------------------------------*/

    /**
     * 将文件写到相应的路径下
     *
     * @param filePath 输出文件路径
     */
    public void write2File(String filePath) throws Excel4JException {

        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            this.workbook.write(fos);
        } catch (IOException e) {
            throw new Excel4JException(e);
        }
    }

    /**
     * 将文件写到某个输出流中
     *
     * @param os 输出流
     */
    public void write2Stream(OutputStream os) throws Excel4JException {

        try {
            this.workbook.write(os);
        } catch (IOException e) {
            throw new Excel4JException(e);
        }
    }

    /*-----------------------------------写出数据结束-----------------------------------*/

    @Override
    public void close() throws IOException {
        if (null != this.workbook){
            this.workbook.close();
        }
    }

}
