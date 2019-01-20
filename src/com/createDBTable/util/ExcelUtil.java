package com.createDBTable.util;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	
	public static Workbook readExcel(String excelFile){
		
		Workbook book = null;
		
			try {
				if (excelFile.toString().toLowerCase().endsWith(".xls")){
					book = new HSSFWorkbook(new FileInputStream(excelFile));
				} else if(excelFile.toString().toLowerCase().endsWith(".xlsx")){
					book = new XSSFWorkbook(new FileInputStream(excelFile));
				}
			} catch (Exception e) {
				e.printStackTrace();
			} 
		return book;
	}
}
