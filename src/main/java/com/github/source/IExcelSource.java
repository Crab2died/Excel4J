package com.github.source;

import org.apache.poi.ss.usermodel.Workbook;

public interface IExcelSource {
	
	/**
	 * 获取源
	 * @return
	 */
	Workbook getWorkBook();
}