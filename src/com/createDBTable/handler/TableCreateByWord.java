package com.createDBTable.handler;

import java.io.BufferedWriter;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.createDBTable.util.ExcelUtil;
import com.createDBTable.util.WordUtil;

public class TableCreateByWord {

	public int rowNum = 3;
	public int cellNum = 1;
	public int idNum = 0;
	public int nameNum = 1;
	public int typeNum = 2;
	public int nullNum = 3;
	public int keyNum = 4;
	public int descNum = 6;
	public int erowNum = 1;
	public int etabEnNum = 0;
	public int etabChNum = 1;
	public int enameNum = 4;
	public int edescNum = 5;
	public int etypeNum = 9;
	public int elenNum = 11;
	public int escaleNum = 12;
	public int enullNum = 14;
	public int ekeyNum = 15;
	public Log log;
	
	public void setLog(Log log){
		this.log = log;
	}
	
	public Connection connectDB(String ip, String port, String sid, String user, String password){
		String driver = "oracle.jdbc.driver.OracleDriver";
		String url = "jdbc:oracle:thin:@"+ip+":"+port+":"+sid;
		try {
			Class.forName(driver);
			Connection conn = DriverManager.getConnection(url, user, password);
			if(conn != null){
				this.log.info("数据库链接成功!!");
				return conn;
			}
			this.log.info("数据库链接失败!!");
		} catch (ClassNotFoundException e) {
			this.log.info("数据库链接失败!!");
			e.printStackTrace();
		} catch (SQLException e) {
			this.log.info("数据库链接失败!!");
			e.printStackTrace();
		} 
		return null;
	}

	public int dealDoc(String wordPath, String sqlPath, String[] dbInfo){
		int rowIndex = 0;
		int cellIndex = 0;
		String tableNameTmp = null;
		String colNameTmp = null;
		String sqlString = null;
		this.log.info("================>开始读取doc文档并生成sql脚本.......");
		try {
			BufferedWriter bfw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(sqlPath), "UTF-8"));
			HWPFDocument hwpf = (new WordUtil()).readByDoc(wordPath);
			Range range = hwpf.getRange();
			TableIterator tbit = new TableIterator(range);
			Table tb = null;
			String nStr = "\n";
			List<String> dropList = new ArrayList<String>();
			List<String> selectList = new ArrayList<String>();
			List<String> createList = new ArrayList<String>();
			List<String> keyList = new ArrayList<String>();
			List<String> commentList = new ArrayList<String>();
			while (tbit.hasNext()){
				tb = tbit.next();
				String tableName = formatStr(tb.getRow(1).getCell(1).text());
				String tableDesc = formatStr(tb.getRow(0).getCell(0).text());
				tableNameTmp = tableName;
				String keyStr = "";
				String selectSql = "";
				String dropSql = "";
				String createSql = "create table ";
				String keySql = "alter table";
				String commentSql = "";
				String WriteCreateSql = "create table ";
				String WriteKeySql = "alter table ";
				String WriteCommentSql = "";
				String sStr = "";
				for (int i=this.rowNum; i < tb.numRows(); i ++){
					rowIndex = i;
					cellIndex = this.nameNum;
					String colName = formatStr(tb.getRow(i).getCell(this.nameNum).text());
					colNameTmp = colName;
					cellIndex = this.typeNum;
					String colType = formatStr(tb.getRow(i).getCell(this.typeNum).text());
					cellIndex = this.nullNum;
					String colNull = formatStr(tb.getRow(i).getCell(this.nullNum).text());
					cellIndex = this.keyNum;
					String colKey = formatStr(tb.getRow(i).getCell(this.keyNum).text());
					cellIndex = this.descNum;
					String colDesc = formatStr(tb.getRow(i).getCell(this.descNum).text().trim().replace("\r", "").replace("\n", "").replace(" ", ""));
					sStr = i < tb.numRows() -1? ",":"";
					if(i==this.rowNum){
						WriteCreateSql = "drop table "+tableName+" cascade constraints;"+nStr+createSql + tableName + nStr + "(" + nStr;
						WriteKeySql = keySql + tableName + " add constraint PK_" + tableName + " primary key (";
						WriteCommentSql = "comment on table " + tableName + "\n is '" + tableDesc + "';" + nStr;
						
						selectSql = "select table_name from user_tables where table_name = '" + tableName + "'";
						dropSql = "drop table " + tableName + " cascade constraints";
						createSql = createSql + tableName + " (";
						keySql = keySql + tableName + " add constraint PK_" + tableName + " primary key (";
						commentSql = "comment on table " + tableName + " is '" + tableDesc +"';";
					}
					WriteCreateSql = WriteCreateSql +" " + colName + " " + colType + (colNull.equalsIgnoreCase("否")?" not null ":"") + sStr + nStr;
					WriteCommentSql = WriteCommentSql + ("".equalsIgnoreCase(colDesc)?"":new StringBuilder("comment on colum ").append(tableName).append(".").append("\n is '").append(colDesc).append("';"));
					createSql = createSql + colName + " " + colType + (colNull.equalsIgnoreCase("否")?" not null":"" + sStr);
					keySql = keyStr + (("主码".equalsIgnoreCase(colKey) || "主键".equalsIgnoreCase(colKey))?colName + sStr:"");
					commentSql = commentSql + ("".equalsIgnoreCase(colDesc)? "" : new StringBuilder("comment on column ").append(tableName).append(colName).append(" is '").append(colDesc).append("';"));
				}
				keyStr = keyStr.length() > 1?keyStr.substring(0, keyStr.length()-1):"";
				WriteCreateSql = WriteCreateSql + ");\n\n";
				WriteKeySql = WriteKeySql + keyStr + ");\n\n";
				WriteCommentSql = WriteCommentSql + "\n";
				createSql = createSql + ")";
				keySql = keySql + keyStr + ")";
				
				bfw.write(WriteCreateSql);
				bfw.write(WriteCommentSql);
				bfw.write(WriteKeySql);
				bfw.flush();
				
				selectList.add(selectSql);
				dropList.add(dropSql);
				createList.add(createSql);
				keyList.add(keySql);
				commentList.add(commentSql);
				
				this.log.info("建表语句："+createSql);
				this.log.info("设置主键："+("".equalsIgnoreCase(keySql)?"无":keySql));
				this.log.info("设置备注："+commentSql);
			}
			bfw.close();
			bfw = null;
			tableNameTmp = null;
			this.log.info("===================>sql脚本生成成功！");
			if(dbInfo!=null){
				this.log.info("===============>开始尝试连接数据库同步建立表结构");
				Connection conn= connectDB(dbInfo[0],dbInfo[1],dbInfo[2],dbInfo[3],dbInfo[4]);
				if(conn !=null){
					Statement stmt = conn.createStatement();
					ResultSet rs = null;
					for(int i=0;i<dropList.size();i++){
						this.log.info("开始建立第"+(i+1)+"张表");
						this.log.info("建立表结构");
						rs = stmt.executeQuery((String)selectList.get(i));
						if(rs.next()){
							sqlString = (String) dropList.get(i);
							stmt.executeUpdate((String)dropList.get(i));
						}
						sqlString = (String)createList.get(i);
						stmt.executeUpdate((String)createList.get(i));
						this.log.info("设置主键");
						sqlString = (String) keyList.get(i);
						if(!"".equalsIgnoreCase(sqlString)){
							stmt.executeUpdate(sqlString);
						}
						this.log.info("设置字段备注");
						String[] commentStr = commentList.get(i).split(";");
						for(int j=0;j<commentStr.length; j++){
							if((commentStr[j])==null || "".equalsIgnoreCase(commentStr[j])){
								continue;
							}
							sqlString = commentStr[j];
							stmt.executeUpdate(commentStr[j]);
						}
						this.log.info("第"+(i+1)+"张表建立成功");
					}
					sqlString = null;
					conn.close();
					this.log.info("所有表结构建立完成");
				}
			}
		} catch (Exception e) {
			if(tableNameTmp != null){
				this.log.info("出错的表名称："+tableNameTmp + ": 字段名："+colNameTmp);
				this.log.info("出错的位置： 编号"+(rowIndex - this.rowNum+1)+"," +(cellIndex)+"列，请检查！");
			}
			if(sqlString != null){
				this.log.info("执行sql语句出错！");
				this.log.info(sqlString);
			}
			this.log.info("文档处理失败！");
			System.gc();
			return -1;
		}
		return 0;
	}
	
	public int dealDocx(String wordPath, String sqlPath, String[] dbInfo){
		int rowIndex = 0;
		int cellIndex = 0;
		String tableNameTmp=null;
		String colNameTmp = null;
		String sqlString = null;
		this.log.info("=========>开始读取Docx文档并生成sql脚本_");
		try{
			BufferedWriter bfw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(sqlPath), "UTF-8"));
			XWPFDocument xwpf = WordUtil.readByDocx(wordPath);
			List tbit = xwpf.getTables();
			XWPFTable tb = null;
			String nStr ="\n";
			List<String> dropList = new ArrayList<String>();
			List<String> selectList = new ArrayList<String>();
			List<String> createList = new ArrayList<String>();
			List<String> keyList = new ArrayList<String>();
			List<String> commentList = new ArrayList<String>();
			for(int i=0;i<tbit.size();i++){
				tb = (XWPFTable)tbit.get(i);
				String tableName = formatStr(tb.getRow(1).getCell(1).getText());
				String tableDesc = formatStr(tb.getRow(0).getCell(1).getText());
				String keyStr ="";
				String selectSql="";
				String dropSql="";
				String createSql="create table ";
				String keySql = "alter table ";
				String commentSql="";
				String WriteCreateSql="create table ";
				String WriteKeySql="alter table ";
				String WriteCommentSql="";
				String sStr ="";
				List tbr = tb.getRows();
				for(int j=this.rowNum;j<tbr.size();j++){
					rowIndex=j;
					cellIndex=this.nameNum;
					String colName = formatStr(((XWPFTableRow)tbr.get(j)).getCell(this.nameNum).getText());
					colNameTmp = colName;
					cellIndex=this.typeNum;
					String colType=formatStr(((XWPFTableRow)tbr.get(j)).getCell(this.typeNum).getText());
					cellIndex=this.nullNum;
					String colNull=formatStr(((XWPFTableRow)tbr.get(j)).getCell(this.nullNum).getText());
					cellIndex=this.keyNum;
					String colKey=formatStr(((XWPFTableRow)tbr.get(j)).getCell(this.keyNum).getText());
					cellIndex=this.descNum;
					String colDesc=formatStr(((XWPFTableRow)tbr.get(j)).getCell(this.descNum).getText());
					sStr = j<tbr.size()-1?",":"";
					if(j==this.rowNum){
						WriteCreateSql="drop table "+tableName+" cascade constraints;"+nStr+createSql+tableName+nStr+"("+nStr;
						WriteKeySql=keySql+tableName+" add constraint PK_"+tableName+" primary key (";
						WriteCommentSql="comment on table "+tableName+"\n is '"+tableDesc+"';"+nStr;
						
						selectSql="select table_name from user_tables where table_name='"+tableName+"'";
						dropSql="drop table "+tableName+" cascade constraints";
						createSql=createSql+tableName+" (";
						keySql=keySql+tableName+" add constraint PK_"+tableName+" primary key (";
						commentSql="comment on table "+tableName+" is '"+tableDesc+"';";
					}
					WriteCreateSql = WriteCreateSql+"  "+colName+"  "+colType+(colNull.equalsIgnoreCase("否")?" not null":"")+sStr+nStr;
					WriteCommentSql=WriteCommentSql+("".equalsIgnoreCase(colDesc)?"":new StringBuilder("comment on column ").append(tableName).append(".").append(colName).append("\n is '").append(colDesc).append("';").append(nStr).toString());
					createSql = createSql+colName+"  "+colType+(colNull.equalsIgnoreCase("否")?" not null":"")+sStr;
					keyStr = keyStr+("主码".equalsIgnoreCase(colKey)||"主码".equalsIgnoreCase(colKey)?colName+sStr:"");
					commentSql=commentSql+("".equalsIgnoreCase(colDesc)?"":new StringBuilder("comment on column ").append(tableName).append(".").append(colName).append(" is '").append(colDesc).append("';").toString());
				}
				keyStr=keyStr.length()>1?keyStr.substring(0, keyStr.length()-1):"";
				WriteCreateSql=WriteCreateSql+");\n\n";
				WriteKeySql=WriteKeySql+keyStr+");\n\n";
				WriteCommentSql=WriteCommentSql+"\n";
				createSql=createSql+")";
				keySql=keySql+keyStr+")";
				bfw.write(WriteCreateSql);
				bfw.write(WriteCommentSql);
				bfw.write(WriteKeySql);
				bfw.flush();
				
				selectList.add(selectSql);
				dropList.add(dropSql);
				createList.add(createSql);
				keyList.add(keySql);
				commentList.add(commentSql);
				this.log.info("建表语句："+createSql);
				this.log.info("设置主键："+("".equalsIgnoreCase(keySql)?"无":keySql));
				this.log.info("设置备注："+commentSql);
			}
			bfw.close();
			bfw=null;
			tableNameTmp=null;
			this.log.info("============>生成sql脚本成功__");
			if(dbInfo != null){
				this.log.info("============>开始尝试连接数据库同步建立表结构___");
				Connection conn = connectDB(dbInfo[0],dbInfo[1],dbInfo[2],dbInfo[3],dbInfo[4]);
				if(conn != null){
					Statement stmt = conn.createStatement();
					ResultSet rs = null;
					for(int i=0;i<dropList.size();i++){
						this.log.info("开始建立第"+(i+1)+"张表");
						this.log.info("建立表结构");
						rs = stmt.executeQuery((String)selectList.get(i));
						if(rs.next()){
							sqlString=dropList.get(i);
							stmt.executeUpdate((String)dropList.get(i));
						}
						sqlString = (String) createList.get(i);
						stmt.executeUpdate((String)createList.get(i));
						this.log.info("设置主键");
						sqlString = (String) keyList.get(i);
						if(!"".equalsIgnoreCase(sqlString)){
							stmt.executeUpdate(sqlString);
						}
						this.log.info("设置字段备注");
						String[] commentStr = ((String) commentList.get(i)).split(";");
						for(int j=0;j<commentStr.length;j++){
							sqlString = commentStr[j];
							stmt.executeUpdate(sqlString);
						}
						this.log.info("第 "+(i+1)+"张表建立成功");
					}
					sqlString = null;
					conn.close();
					this.log.info("所有表结构建立完成");
				}
			}
		}catch(Exception e){
			if(tableNameTmp != null){
				this.log.info("出错的表名称："+tableNameTmp + ": 字段名："+colNameTmp);
				this.log.info("出错的位置： 编号"+(rowIndex - this.rowNum+1)+"," +(cellIndex)+"列，请检查！");
			}
			if(sqlString != null){
				this.log.info("执行sql语句出错！");
				this.log.info(sqlString);
			}
			this.log.info("文档处理失败！");
			System.gc();
			return -1;
		}
		return 0;
	}
	
	public int dealExcel(String excelPath, String sqlPath, String[] dbInfo, int dealType){
		int rowIndex = 0;
		int cellIndex = 0;
		String tableNameTmp = null;
		String colNameTmp = null;
		String sqlString = null;
		this.log.info("=======>开始读取excel文档并生成SQL脚本___");
		try{
			BufferedWriter bfw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(sqlPath),"UTF-8"));
			Workbook table = ExcelUtil.readExcel(excelPath);
			Sheet sheet = table.getSheetAt(0);
			String nStr = "\n";
			String tableName;
			String t2;
			String t1 = t2 = tableName ="";
			List<String> dropList = new ArrayList<String>();
			List<String> selectList = new ArrayList<String>();
			List<String> createList = new ArrayList<String>();
			List<String> keyList = new ArrayList<String>();
			List<String> commentList = new ArrayList<String>();
			int i=0;
			while(i<=sheet.getLastRowNum()){
				if(t1.equalsIgnoreCase("")){
					if(null == sheet.getRow(i) || sheet.getRow(i).getCell(this.etabEnNum).toString().trim().equalsIgnoreCase("")){
						i++;
						continue;
					}
					t1=formatStr(sheet.getRow(i).getCell(this.etabEnNum).toString().trim());
					i++;
				} else {
					t1=tableName;
				}
				if(null == sheet.getRow(i) || sheet.getRow(i).getCell(this.etabEnNum).toString().trim().equalsIgnoreCase("")){
					i++;
				} else {
					t2=formatStr(sheet.getRow(i).getCell(this.etabEnNum).toString().trim());
					this.erowNum=1;
					if(t1.equalsIgnoreCase(t2)){
						continue;
					}
					tableNameTmp = tableName = t2;
					String tableDesc = formatStr(sheet.getRow(i).getCell(this.etabChNum).toString().trim());
					String keyStr="";
					String selectSql="";
					String dropSql="";
					String createSql="create table ";
					String keySql="alter table ";
					String commentSql="";
					String WriteCreateSql="create table ";
					String WriteKeySql="alter table ";
					String WriteCommentSql="";
					String sStr = "";
					while(i<=sheet.getLastRowNum()){
						Row row = sheet.getRow(i);
						if(null == row || null == row.getCell(this.etabEnNum) || row.getCell(this.etabEnNum).toString().trim().equalsIgnoreCase("")){
							i++;
						} else {
							if(!formatStr(row.getCell(this.etabEnNum).toString().trim()).equalsIgnoreCase(tableName)){
								break;
							}
							setCellType(row);
							rowIndex=i;
							cellIndex = this.enameNum;
							String colName = formatStr(row.getCell(this.enameNum).toString().trim());
							colNameTmp = colName;
							cellIndex = this.etypeNum;
							String colType = formatStr(row.getCell(this.etypeNum).toString().trim());
							cellIndex = this.elenNum;
							String colLen = formatStr(row.getCell(this.elenNum).toString().trim());
							cellIndex = this.escaleNum;
							String colScale = formatStr(row.getCell(this.escaleNum).toString().trim());
							cellIndex = this.enullNum;
							String colNull = formatStr(row.getCell(this.enullNum).toString().trim());
							cellIndex = this.ekeyNum;
							String colKey = formatStr(row.getCell(this.ekeyNum).toString().trim());
							cellIndex = this.edescNum;
							String colDesc = formatStr(row.getCell(this.edescNum).toString().trim());
							sStr = i<sheet.getLastRowNum()?",":"";
							if(i == this.erowNum){
								WriteCreateSql = "drop table "+tableName +" cascade constarints;"+nStr+createSql+tableName+nStr+"("+nStr;
								WriteKeySql=keySql+tableName+" add constraint PK_"+tableName+" primary key (";
								WriteCommentSql="comment on table "+tableName+"\n is '"+tableDesc +"';"+nStr;
								
								selectSql = "select table_name from user_tables where table_name='"+tableName+"'";
								dropSql="drop table "+tableName+" cascade constraints";
								createSql=createSql+tableName+" (";
								keySql = keySql+tableName+" add constraint PK_"+tableName+" primary key (";
								commentSql="comment on table "+tableName+" is '"+tableDesc+"';";
							}
							WriteCreateSql = WriteCreateSql + " " + colName + " " + (colType.equalsIgnoreCase("DATE")
									? colType
									: dealType == 0 ? "VARCHER2(" + colLen + ")"
											: new StringBuilder(String.valueOf(colType)).append("(").append(colLen)
													.append((!colScale.equalsIgnoreCase(""))
															&& (!colScale.equalsIgnoreCase("0")) ? "," + colScale : "")
													.append(")").toString())
									+ (colNull.equalsIgnoreCase("N") ? " not null" : "") + sStr + nStr;
							WriteCommentSql = WriteCommentSql + ("".equalsIgnoreCase(colDesc) ? ""
									: new StringBuilder("comment on column ").append(tableName).append(".")
											.append(colName).append("\n is '").append(colDesc).append(""));//待确认
							createSql = createSql + colName + " " + (colType.equalsIgnoreCase("DATE") ? colType
									: dealType == 0 ? "VARCHER2(" + colLen + ")"
											: new StringBuilder(String.valueOf(colType)).append("(").append(colLen)
													.append((!colScale.equalsIgnoreCase(""))
															&& (!colScale.equalsIgnoreCase("0")) ? "," + colScale : "")
													.append(")").toString())
									+ (colNull.equalsIgnoreCase("N") ? " not null" : "") + sStr;
							keyStr = keyStr+("Y".equalsIgnoreCase(colKey)?colName+sStr:"");
							commentSql = commentSql + ("".equalsIgnoreCase(colDesc) ? ""
									: new StringBuilder("comment on column ").append(tableName).append(".")
											.append(colName).append(" is '").append(colDesc).append(""));// 待确定
							i++;
						}
					}
					keyStr = keyStr.length()>1?keyStr.substring(0, keyStr.length()-1):"";
					WriteCreateSql = (i < sheet.getLastRowNum()
							? WriteCreateSql.substring(0, WriteCreateSql.length() - 2) + "\n" : WriteCreateSql) + "";// 待确定
					WriteKeySql = WriteKeySql+keyStr+");\n\n";
					WriteCommentSql=WriteCommentSql+"\n";
					createSql=(i < sheet.getLastRowNum()
							? createSql.substring(0, createSql.length() - 1):createSql)+")";
					keySql = keySql+keyStr+")";
					
					bfw.write(WriteCreateSql);
					bfw.write(WriteCommentSql);
					bfw.write(WriteKeySql);
					bfw.flush();
					
					selectList.add(selectSql);
					dropList.add(dropSql);
					createList.add(createSql);
					keyList.add(keySql);
					commentList.add(commentSql);
					this.log.info("建表语句："+createSql);
					this.log.info("设置主键："+("".equalsIgnoreCase(keySql)?"无":keySql));
					this.log.info("设置备注："+commentSql);
				}
			}
			
			bfw.close();
			bfw=null;
			tableNameTmp=null;
			this.log.info("============>生成sql脚本成功__");
			if(dbInfo != null){
				this.log.info("============>开始尝试连接数据库同步建立表结构___");
				Connection conn = connectDB(dbInfo[0],dbInfo[1],dbInfo[2],dbInfo[3],dbInfo[4]);
				if(conn != null){
					Statement stmt = conn.createStatement();
					ResultSet rs = null;
					for(int j=0;j<dropList.size();j++){
						this.log.info("开始建立第"+(j+1)+"张表");
						this.log.info("建立表结构");
						rs = stmt.executeQuery((String)selectList.get(j));
						if(rs.next()){
							sqlString=dropList.get(j);
							stmt.executeUpdate((String)dropList.get(j));
						}
						sqlString = (String) createList.get(j);
						stmt.executeUpdate((String)createList.get(j));
						this.log.info("设置主键");
						sqlString = (String) keyList.get(j);
						if(!"".equalsIgnoreCase(sqlString)){
							stmt.executeUpdate(sqlString);
						}
						this.log.info("设置字段备注");
						String[] commentStr = ((String) commentList.get(j)).split(";");
						for(int m=0;m<commentStr.length;m++){
							sqlString = commentStr[m];
							stmt.executeUpdate(sqlString);
						}
						this.log.info("第 "+(j+1)+"张表建立成功");
					}
					sqlString = null;
					conn.close();
					this.log.info("所有表结构建立完成");
				}
			}
		}catch(Exception e){
			if(tableNameTmp != null){
				this.log.info("出错的表名称："+tableNameTmp + ": 字段名："+colNameTmp);
				this.log.info("出错的位置： 编号"+(rowIndex - this.rowNum+1)+"," +(cellIndex)+"列，请检查！");
			}
			if(sqlString != null){
				this.log.info("执行sql语句出错！");
				this.log.info(sqlString);
			}
			this.log.info("文档处理失败！");
			System.gc();
			return -1;
		}
		return 0;
	}
	
	public String formatStr(String str){
		return str.trim().replace("\r", "").replace("\n", "").replace(" ", "");
	}
	
	public String getFirstLine(String str){
		int n = str.indexOf("\n");
		int r = str.indexOf("\r");
		if(n<0 && r<0){
			return str;
		}
		if(n<r && n>0){
			return str.substring(0, n);
		}
		if(r<n && r>0){
			return str.substring(0, r);
		}
		if(n>0 && r<0){
			return str.substring(0, n);
		}
		if(r>0 && r<0){
			return str.substring(0, r);
		}
		return str;
	}
	
	public String getNowTimes(String format){
		Date date = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat(format);
		return sdf.format(date);
	}
	
	public void setCellType(Row row) {
		if(null != row){
			for(int m=0;m<row.getLastCellNum();m++){
				if(null == row.getCell(m)){
					continue;
				}
				row.getCell(m).setCellType(1);
			}
		}
	}
	
	public static void main(String[] args){
		try{
			TableCreateByWord tc = new TableCreateByWord();
			String wordPath = "";
			String sqlPath = "";
			
			String[] dbInfo = null;
			tc.dealExcel(wordPath, sqlPath, dbInfo, 0);
		}catch(Exception e){
			e.printStackTrace();
		}
	}
}
