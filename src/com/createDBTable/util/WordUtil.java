package com.createDBTable.util;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class WordUtil {

	public static HWPFDocument readByDoc(String filePath){
		HWPFDocument hwpf = null;
		try {
			hwpf = new HWPFDocument(new FileInputStream(filePath));
		} catch (Exception e) {
			e.printStackTrace();
		}
		return hwpf;
	}
	
	public static XWPFDocument readByDocx(String filePath){
		XWPFDocument xwpf = null;
		try {
			xwpf = new XWPFDocument(new FileInputStream(filePath));
		} catch (Exception e) {
			e.printStackTrace();
		}
		return xwpf;
	}
}
