package com.github.sink;

import java.io.OutputStream;

public interface IExcelSink {
	
	/**
	 * 获取到输出的OutputStream
	 * @return
	 */
	OutputStream getSink();
	
	/**
	 * 关闭Sink
	 */
	void close();
}