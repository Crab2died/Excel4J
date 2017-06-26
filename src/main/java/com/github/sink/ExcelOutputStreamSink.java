package com.github.sink;

import java.io.IOException;
import java.io.OutputStream;

public class ExcelOutputStreamSink implements IExcelSink{
	
	private OutputStream outputStream;
	
	/**
	 * 创建OutputStreamSink
	 * @param outputStream
	 * @return
	 */
	public static ExcelOutputStreamSink create(OutputStream outputStream) {
		ExcelOutputStreamSink excelFileSink = new ExcelOutputStreamSink();
		excelFileSink.outputStream = outputStream;
		return excelFileSink;
	}

	@Override
	public OutputStream getSink() {
		return outputStream;
	}

	@Override
	public void close() {
		try {
            if (outputStream != null)
            	outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
	}
}