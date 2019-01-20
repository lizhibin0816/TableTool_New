package com.createDBTable.util;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;
import java.util.Set;

public class Config {

	public Properties prop = new Properties();
	public String fileName = null;
	public boolean nullprop = true;
	
	public Config(String fileName){
		this.fileName = fileName;
		initConfig(fileName);
	}
	
	public void initConfig(String fileName){
		String filePath = System.getProperty("user.dir")+"\\"+fileName;
		File file = new File(filePath);
		if(file.exists()){
			try {
				BufferedInputStream bins = new BufferedInputStream(new FileInputStream(file));
				this.prop.load(bins);
				bins.close();
				this.nullprop = false;
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	public String getValue(String key){
		return this.prop.getProperty(key);
	}
	
	public void setValue(String key, String value){
		this.prop.setProperty(key, value);
	}
	
	public void setValueByMap(Map<String,String> map){
		Set<String> keySet = map.keySet();
		Iterator<String> keyIt = keySet.iterator();
		while(keyIt.hasNext()){
			String key = keyIt.next();
			this.prop.setProperty(key, map.get(key));
		}
	}
}
