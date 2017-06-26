package com.github.source;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileSource implements IExcelSource {
	
	private Workbook workbook = null;
		
	private ExcelFileSource(){}
	
	/**
	 * 创建一个文本Excel源
	 * @return
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 * @throws EncryptedDocumentException 
	 */
	public static IExcelSource create(String excelPath) throws Exception{
		ExcelFileSource excelFileSource = new ExcelFileSource();
		excelFileSource.workbook = WorkbookFactory.create(new File(excelPath));
		
		return excelFileSource;
	}

	@Override
	public Workbook getWorkBook() {
		return workbook;
	}
}
