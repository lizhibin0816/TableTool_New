package com.createDBTable.log;

import java.util.Scanner;

import javax.swing.JScrollPane;
import javax.swing.JTextArea;

public class TextAreaLogAppender extends LogAppender {

	private JTextArea textArea;
	private JScrollPane scroll;
	
	public TextAreaLogAppender(JTextArea textArea, JScrollPane scroll){
		super("textArea");
		this.textArea = textArea;
		this.scroll = scroll;
	}
	
	@Override
	public void run(){
		Scanner scanner = new Scanner(this.reader);
		while(scanner.hasNextLine()){
			try {
				Thread.sleep(100L);
				String line = scanner.nextLine();
				this.textArea.append(line);
				this.textArea.append("\n");
				line = null;
				
				this.scroll.getVerticalScrollBar().setValue(this.scroll.getVerticalScrollBar().getMaximum());
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
	}
}
