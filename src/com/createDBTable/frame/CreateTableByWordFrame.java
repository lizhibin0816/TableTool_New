package com.createDBTable.frame;

import java.awt.Component;
import java.awt.Container;
import java.awt.Point;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.io.File;
import java.util.HashMap;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import com.createDBTable.handler.TableCreateByWord;
import com.createDBTable.log.TextAreaLogAppender;
import com.createDBTable.util.Config;

public class CreateTableByWordFrame implements ActionListener, ItemListener {

	public static Log log = LogFactory.getLog(CreateTableByWordFrame.class);
	JFrame frame = new JFrame("建表脚本生成工具");
	JLabel wordLabel = new JLabel("表结构路径");
	JLabel sqlLabel = new JLabel("目标脚本路径");
	JLabel ipLabel = new JLabel("数据库ip");
	JLabel portLabel = new JLabel("数据库端口");
	JLabel sidLabel =  new JLabel("数据库sid");
	JLabel userLabel  = new JLabel("数据库用户");
	JLabel pswLabel = new JLabel("数据库密码");
	
	JTextField wordPathField = new JTextField();
	JTextField sqlPathField = new JTextField();
	JTextField ipField = new JTextField();
	JTextField portField = new JTextField();
	JTextField sidField =  new JTextField();
	JTextField userField  = new JTextField();
	JTextField pswField = new JTextField();
	
	JButton selectWordButton = new JButton("选择");
	JButton selectSqlButton = new JButton("选择");
	JButton creButton = new JButton("生成sql脚本");
	
	Container container = new Container();
	JFileChooser chooser = new JFileChooser();
	
	double screenHight = Toolkit.getDefaultToolkit().getScreenSize().getHeight();
	double screenWidth = Toolkit.getDefaultToolkit().getScreenSize().getWidth();
	
	String selectWordPath = null;
	String selectSqlPath = null;
	String sqlName = "\\Create.sql";
	String defaultDbSelect = null;
	String defaultIp = null;
	String defaultPort = null;
	String defaultSid = null;
	String defaultUser = null;
	String defaultPsw = null;
	String defaultSqlMode = null;
	
	JTextArea logTextArea = new JTextArea();
	JScrollPane logPane = new JScrollPane(this.logTextArea);
	
	JCheckBox cb = new JCheckBox("同时连接数据库创建表");
	
	public Config config = new Config("app.properties");
	String docKey = "app.TableCreateByWord.docPath";
	String SqlKey = "app.TableCreateByWord.sqlPath";
	String dbSelectKey = "app.TableCreateByWord.dbIsSelect";
	String dbIpKey = "app.TableCreateByWord.dbIp";
	String dbPortKey = "app.TableCreateByWord.dbPort";
	String dbSidKey = "app.TableCreateByWord.dbSid";
	String dbUserKey = "app.TableCreateByWord.dbUser";
	String dbPswKey = "app.TableCreateByWord.dbPsw";
	String dbSqlModeKey = "app.TableCreateByWord.sqlMode";
	
	public TableCreateByWord tcbw = null;
	public String wordPath = "";
	public String sqlPath = "";
	public String[] dbInfo = null;
	
	public static void main(String[] args){
		new CreateTableByWordFrame();
	}
	
	public void initTCBW(){
		initLog();
		this.tcbw = new TableCreateByWord();
		this.tcbw = setLog(log);
	}
	
	public void initLog(){
		try{
			Thread logThread = new TextAreaLogAppender(this.logTextArea, this.logPane);
			this.logTextArea.append(" ");
			logThread.start();
		} catch (Exception e) {
			JOptionPane.showMessageDialog(null, "日志组件出错！","错误提示",0);
		}
	}
	
	public CreateTableByWordFrame(){
		this.selectWordPath = this.config.getValue(this.docKey);
		this.selectSqlPath = this.config.getValue(this.SqlKey);
		this.defaultPort = this.config.getValue(dbPortKey);
		this.defaultSid = this.config.getValue(dbSidKey);
		this.defaultIp = this.config.getValue(dbIpKey);
		this.defaultDbSelect = this.config.getValue(dbSelectKey);
		this.defaultUser = this.config.getValue(dbUserKey);
		this.defaultPsw = this.config.getValue(dbPswKey);
		this.defaultSqlMode = this.config.getValue(dbSqlModeKey);
		if("".equalsIgnoreCase(this.defaultPort) || this.defaultPort == null){
			this.defaultPort = "1521";
		}
		if("".equalsIgnoreCase(this.defaultSid) || this.defaultSid == null){
			this.defaultPort = "orau11g";
		}
		
		this.wordLabel.setBounds(60, 30, 80, 20);
		this.sqlLabel.setBounds(60, 70, 80, 20);
		this.ipLabel.setBounds(60, 110, 80, 20);
		this.cb.setBounds(400, 98, 170, 40);
		this.sidLabel.setBounds(60, 150, 80, 20);
		this.portLabel.setBounds(320, 150, 80, 20);
		this.userLabel.setBounds(60, 190, 80, 20);
		this.pswLabel.setBounds(320, 190, 80, 20);
		
		this.wordPathField.setBounds(150, 30, 300, 20);
		this.sqlPathField.setBounds(150, 70, 300, 20);
		this.ipField.setBounds(150, 110, 140, 20);
		this.portField.setBounds(410, 150, 140, 20);
		this.sidField.setBounds(150, 150, 140, 20);
		this.userField.setBounds(150, 190, 140, 20);
		this.pswField.setBounds(410, 190, 140, 20);
		this.selectWordButton.setBounds(470, 30, 80, 20);
		this.selectSqlButton.setBounds(470, 70, 80, 20);
		this.creButton.setBounds(250, 400, 120, 20);
		
		this.selectWordButton.addActionListener(this);
		this.selectSqlButton.addActionListener(this);
		this.creButton.addActionListener(this);
		this.cb.addItemListener(this);
		
		this.logPane.setBounds(60, 230, 490, 150);
		
		this.wordPathField.setText(this.selectWordPath);
		this.sqlPathField.setText(this.selectSqlPath);
		this.portField.setText(this.defaultPort);;
		this.sidField.setText(this.defaultSid);
		this.ipField.setText(this.defaultIp);
		this.userField.setText(this.defaultUser);
		this.pswField.setText(this.defaultPsw);
		
		this.ipField.setEditable(false);
		this.portField.setEditable(false);
		this.sidField.setEditable(false);
		this.userField.setEditable(false);
		this.pswField.setEditable(false);
		this.cb.setSelected(false);
		
		this.container.add(this.wordLabel);
		this.container.add(this.sqlLabel);
		this.container.add(this.ipLabel);
		this.container.add(this.sidLabel);
		this.container.add(this.portLabel);
		this.container.add(this.userLabel);
		this.container.add(this.pswLabel);
		this.container.add(this.wordPathField);
		this.container.add(this.sqlPathField);
		this.container.add(this.ipField);
		this.container.add(this.portField);
		this.container.add(this.sidField);
		this.container.add(this.userField);
		this.container.add(this.pswField);
		this.container.add(this.selectWordButton);
		this.container.add(this.selectSqlButton);
		this.container.add(this.creButton);
		this.container.add(this.logPane);
		this.container.add(this.cb);
		
		this.frame.setDefaultCloseOperation(3);
		this.frame.add(this.container);
		this.frame.pack();
		this.frame.setSize(640, 500);
		this.frame.setLocation(new Point((int)(this.screenWidth/2.0D)-500,(int)(this.screenHight/2.0D)-350));
		this.frame.setVisible(true);
	}
	
	public void removeComponent(Component com){
		
		if(com != null){
			this.container.remove(com);
		}
	}
	
	public void repaintFrame(Component com){
		
		this.container.add(com);
		this.frame.add(this.container);
		this.frame.repaint();
		this.frame.setVisible(true);
	}
	
	public TableCreateByWord setLog(Log log){
		this.log = log;
		return tcbw;
	}
	
	@Override
	public void actionPerformed(ActionEvent e) {
		if(e.getSource().equals(this.selectWordButton) || e.getSource().equals(this.selectSqlButton)){
			int selectType = 1;
			if(e.getSource().equals(this.selectWordButton)){
				selectType = 0;
			}
			this.chooser.setFileSelectionMode(selectType);
			int state = this.chooser.showOpenDialog(null);
			if(state == 1){
				return;
			}
			File f = this.chooser.getSelectedFile();
			String filePath = f.getAbsolutePath();
			if(e.getSource().equals(this.selectWordButton)){
				if(!filePath.endsWith(".doc") && !filePath.endsWith(".docx") && !filePath.endsWith(".xls") && !filePath.endsWith(".xlsx")){
					JOptionPane.showMessageDialog(null, "只支持文档后缀为doc和docx", "错误提示", 0);
				}
				this.wordPathField.setText(filePath);
				this.config.setValue(this.docKey, filePath);
			} else if (e.getSource().equals(this.selectSqlButton)){
				if (filePath.endsWith("\\")){
					filePath = filePath.substring(0, filePath.length()-1);
				}
				this.sqlPathField.setText(filePath + this.sqlName);
				this.config.setValue(this.SqlKey, filePath+this.sqlName);
			} else if (e.getSource().equals(this.creButton)){
				Map<String, String> m = new HashMap<String, String>();
				this.wordPath = this.wordPathField.getText().trim();
				this.sqlPath = this.sqlPathField.getText().trim();
				if (this.wordPath.length() == 0 || this.sqlPath.length() == 0){
					JOptionPane.showMessageDialog(null, "路径信息不能为空！", "错误提示", 0);
					return;
				}
				this.dbInfo = null;
				if (this.cb.isSelected()){
					String ip = this.ipField.getText().trim();
					String port = this.portField.getText().trim();
					String sid = this.sidField.getText().trim();
					String user = this.userField.getText().trim();
					String password = this.pswField.getText().trim();
					if (ip.length() ==0 || port.length() == 0 || sid.length() == 0 || user.length() == 0 || password.length() == 0){
						JOptionPane.showMessageDialog(null, "数据库信息不能为空！", "错误提示", 0);
						return;
					}
					this.dbInfo = new String[]{ip,port,sid,user,password};
					m.put(this.dbIpKey, ip);
					m.put(this.dbPortKey, port);
					m.put(this.dbSidKey, sid);
					m.put(this.dbUserKey, user);
					m.put(this.dbPswKey, password);
				}
				m.put(this.docKey, this.wordPath);
				m.put(this.SqlKey, this.sqlPath);
				if("".equalsIgnoreCase(this.defaultSqlMode) || null == this.defaultSqlMode){
					m.put(this.dbSqlModeKey, "1");
					this.defaultSqlMode = "1";
				}
				this.config.setValueByMap(m);
				initTCBW();
				if(this.wordPath.endsWith(".doc")){
					new Thread(new Runnable(){
						public void run(){
							CreateTableByWordFrame.this.tcbw.dealDoc(
									CreateTableByWordFrame.this.wordPath, 
									CreateTableByWordFrame.this.sqlPath, 
									CreateTableByWordFrame.this.dbInfo
							);
						}
					}).start();
				} else if (this.wordPath.endsWith(".docx")){
					new Thread(new Runnable(){
						public void run(){
							CreateTableByWordFrame.this.tcbw.dealDocx(
									CreateTableByWordFrame.this.wordPath, 
									CreateTableByWordFrame.this.sqlPath, 
									CreateTableByWordFrame.this.dbInfo 
							);
						}
					}).start();
				} else if (this.wordPath.endsWith(".xls") || this.wordPath.endsWith(".xlsx")){
					new Thread(new Runnable(){
						public void run(){
							CreateTableByWordFrame.this.tcbw.dealExcel(
									CreateTableByWordFrame.this.wordPath, 
									CreateTableByWordFrame.this.sqlPath, 
									CreateTableByWordFrame.this.dbInfo, 
									CreateTableByWordFrame.this.defaultSqlMode.equalsIgnoreCase("0")?0:1
							);
						}
					}).start();
				}
			}
		}
	}
	
	@Override
	public void itemStateChanged(ItemEvent e) {
		JCheckBox jcb = (JCheckBox) e.getItem();
		if(e.getSource().equals(this.cb)){
			if(jcb.isSelected()){
				this.ipField.setEditable(true);
				this.portField.setEditable(true);
				this.sidField.setEditable(true);
				this.userField.setEditable(true);
				this.pswField.setEditable(true);
			} else {
				this.ipField.setEditable(false);
				this.portField.setEditable(false);
				this.sidField.setEditable(false);
				this.userField.setEditable(false);
				this.pswField.setEditable(false);
			}
		}
	}

}
