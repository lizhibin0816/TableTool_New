package com.createDBTable.log;

import java.io.IOException;
import java.io.PipedReader;
import java.io.PipedWriter;
import java.io.Writer;

import org.apache.log4j.Appender;
import org.apache.log4j.Logger;
import org.apache.log4j.WriterAppender;

public abstract class LogAppender extends Thread {

	protected PipedReader reader;
	
	public LogAppender(String appenderName){
		try {
		Logger root = Logger.getRootLogger();
		Appender appender = root.getAppender(appenderName);
		this.reader = new PipedReader();
		Writer writer = new PipedWriter(this.reader);
		((WriterAppender)appender).setWriter(writer);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
