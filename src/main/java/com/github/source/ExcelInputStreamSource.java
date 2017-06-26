package com.github.source;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelInputStreamSource implements IExcelSource{
	
	private Workbook workbook = null;
	
	private ExcelInputStreamSource(){}
	
	
	/**
	 * 创建一个InputStream的Excel源
	 * @return
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 * @throws EncryptedDocumentException 
	 */
	public static IExcelSource create(InputStream inputStream) throws Exception{
		ExcelInputStreamSource excelFileSource = new ExcelInputStreamSource();
		excelFileSource.workbook = WorkbookFactory.create(inputStream);
		
		return excelFileSource;
	}

	@Override
	public Workbook getWorkBook() {
		return workbook;
	}
}