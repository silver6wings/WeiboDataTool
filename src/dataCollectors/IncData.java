package dataCollectors;

import dataTools.ExcelTools;

import java.io.FileInputStream;
import java.net.HttpURLConnection;
import java.net.URL;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.htmlparser.NodeFilter;
import org.htmlparser.Parser;
import org.htmlparser.filters.NodeClassFilter;
import org.htmlparser.tags.TableColumn;
import org.htmlparser.tags.TableRow;
import org.htmlparser.tags.TableTag;
import org.htmlparser.util.NodeList;

public class IncData {

	/*
	 * 日常需要做的工作：
	 * 1.改动main函数的里边的日期
	 * 2.运行程序跑出结果
	 * 3.检查结果是否异常
	 * 4.若没有异常则将结果复制粘贴进增量汇总表并邮件汇报
	 * 5.若有异常一般为发新版丢数据，则执行下边的步骤
	 * 	1）将增量模板的数据添加一大列并标记好新的From值
	 * 	2）在最左列更新橙色部分的公式（多一个求和项）
	 * 	3）切换至汇总界面Ctrl+A全选之后保存
	 * 	4）代码里更新getDataForDay添加对应的getDataByFrom（注意保持列数对应）
	 * 	5）重新执行步骤2来统计数据结果
	 * 6.若异常并非由于发新版丢数据则找林华协调沟通解决
	 */
	
	public static void main(String[] args) {
		// 调整这里
		getDataForDay("02-04", "D:/数据增量/增量模板.xls");
		getDataForDay("02-05", "D:/数据增量/增量模板.xls");
		getDataForDay("02-06", "D:/数据增量/增量模板.xls");
		getDataForDay("02-07", "D:/数据增量/增量模板.xls");
		getDataForDay("02-08", "D:/数据增量/增量模板.xls");
	}
	
	/**
	 * 每天需要统计的数据
	 * @param today 今天的日期
	 * @param samplePath 模板的路径
	 */
	static void getDataForDay(String today, String samplePath) {

		String myDate = "2015-" + today;
		String myFilePath = "D:/数据增量/增量" + today + ".xls";
		
		try {
			HSSFWorkbook myWorkbook = new HSSFWorkbook(new FileInputStream(samplePath));
			getDataByFrom(myDate, "1045595010", myWorkbook, "Android", 1);
			getDataByFrom(myDate, "1046095010", myWorkbook, "Android", 2);
			getDataByFrom(myDate, "1046195010", myWorkbook, "Android", 3);
			getDataByFrom(myDate, "1046295010", myWorkbook, "Android", 4);
			getDataByFrom(myDate, "1050015010", myWorkbook, "Android", 5);
			getDataByFrom(myDate, "1050095010", myWorkbook, "Android", 6);
			getDataByFrom(myDate, "1051095010", myWorkbook, "Android", 7);

			getDataByFrom(myDate, "1046093010", myWorkbook, "iPhone", 1);
			getDataByFrom(myDate, "1046193010", myWorkbook, "iPhone", 2);
			getDataByFrom(myDate, "1046593010", myWorkbook, "iPhone", 3);
			getDataByFrom(myDate, "1046693010", myWorkbook, "iPhone", 4);
			getDataByFrom(myDate, "1050093010", myWorkbook, "iPhone", 5);
			getDataByFrom(myDate, "1050193010", myWorkbook, "iPhone", 6);
			getDataByFrom(myDate, "1050293010", myWorkbook, "iPhone", 7);
			getDataByFrom(myDate, "1051093010", myWorkbook, "iPhone", 8);
			
			for (int i = 0; i < 3; i++) {
				myWorkbook.getSheetAt(i).setForceFormulaRecalculation(true);
			}
			System.out.println("Outing...");
			ExcelTools.workbookOut(myWorkbook, myFilePath);
			System.out.println("Done!");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 将网页上的数据取下塞入模板
	 * @param date 日期
	 * @param from 版本号
	 * @param wb 目标存储文件
	 * @param sheet 目标表名
	 * @param c 表的第几大列（不要写0为源数据页）
	 * @throws Exception
	 */
	static void getDataByFrom(String date, String from, HSSFWorkbook wb, String sheet, int c) throws Exception {
		int w = 8;
		String strURL = "http://172.16.193.178/utils/incr_stats2?date=" + date + "&from=" + from;

		Parser parser = new Parser((HttpURLConnection) (new URL(strURL)).openConnection());
		parser.setEncoding("utf-8");

		NodeFilter filter = new NodeClassFilter(TableTag.class);
		NodeList nodeList = parser.parse(filter);
		
		ExcelTools.writeStrTo(date, wb, sheet, 0, w * c);
		ExcelTools.writeStrTo(from, wb, sheet, 1, w * c);
		
		for (int i = 0; i < nodeList.size(); i++) {
			if (nodeList.elementAt(i) instanceof TableTag) {
				TableTag tag = (TableTag) nodeList.elementAt(i);
				TableRow[] rows = tag.getRows();

				if (i==0){
					for (int j = 0; j < rows.length; j++) {
						TableRow row = (TableRow) rows[j];
						TableColumn[] columns = row.getColumns();
						for (int k = 0; k < columns.length; ++k) {
							String info = columns[k].toPlainTextString().trim();
							System.out.println(info);

							ExcelTools.writeNumTo(parseContent(info), wb, sheet, 5 + j, w * c + k + 1);
						}
					}
				}
				
				if(i==1){
					for (int j = 0; j < rows.length; j++) {
						TableRow row = (TableRow) rows[j];
						TableColumn[] columns = row.getColumns();
						for (int k = 0; k < columns.length; ++k) {
							String info = columns[k].toPlainTextString().trim();
							System.out.println(info);

							ExcelTools.writeNumTo(parseContent(info), wb, sheet, 10, w * c + k);
						}
					}
				}
				
				if (i==2){
					for (int j = 0; j < rows.length; j++) {
						TableRow row = (TableRow) rows[j];
						TableColumn[] columns = row.getColumns();
						for (int k = 0; k < columns.length; ++k) {
							String info = columns[k].toPlainTextString().trim();
							System.out.println(info);

							ExcelTools.writeNumTo(parseContent(info), wb, sheet, 12 + j, w * c + k + 1);
						}
					}
				}
				
				if (i==3){
					for (int j = 0; j < rows.length-1; j++) {
						TableRow row = (TableRow) rows[j];
						TableColumn[] columns = row.getColumns();
						for (int k = 0; k < columns.length; ++k) {
							String info = columns[k].toPlainTextString().trim();
							System.out.println(info);
							ExcelTools.writeNumTo(parseContent(info), wb, sheet, 19 + j, w * c + k + 1);
						}
					}

					int j = rows.length-1;
					TableRow row = (TableRow) rows[j];
					TableColumn[] columns = row.getColumns();
					for (int k = 0; k < columns.length; ++k) {
						String info = columns[k].toPlainTextString().trim();
						System.out.println(info);
						ExcelTools.writeNumTo(parseContent(info), wb, sheet, 70, w * c + k + 1);
					}
				}
			}
		}
	}

	public static float parseContent(String str) {
		if (str.equals("-"))
			return 0;

		if (str.contains("%")) {
			str = str.replace("%", "");
			return Float.parseFloat(str) / 100;
		}

		str = str.replace(",", "");
		return Float.parseFloat(str);
	}
}
