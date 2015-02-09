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
	 * �ճ���Ҫ���Ĺ�����
	 * 1.�Ķ�main��������ߵ�����
	 * 2.���г����ܳ����
	 * 3.������Ƿ��쳣
	 * 4.��û���쳣�򽫽������ճ�����������ܱ��ʼ��㱨
	 * 5.�����쳣һ��Ϊ���°涪���ݣ���ִ���±ߵĲ���
	 * 	1��������ģ����������һ���в���Ǻ��µ�Fromֵ
	 * 	2���������и��³�ɫ���ֵĹ�ʽ����һ������
	 * 	3���л������ܽ���Ctrl+Aȫѡ֮�󱣴�
	 * 	4�����������getDataForDay��Ӷ�Ӧ��getDataByFrom��ע�Ᵽ��������Ӧ��
	 * 	5������ִ�в���2��ͳ�����ݽ��
	 * 6.���쳣�������ڷ��°涪���������ֻ�Э����ͨ���
	 */
	
	public static void main(String[] args) {
		// ��������
		getDataForDay("02-04", "D:/��������/����ģ��.xls");
		getDataForDay("02-05", "D:/��������/����ģ��.xls");
		getDataForDay("02-06", "D:/��������/����ģ��.xls");
		getDataForDay("02-07", "D:/��������/����ģ��.xls");
		getDataForDay("02-08", "D:/��������/����ģ��.xls");
	}
	
	/**
	 * ÿ����Ҫͳ�Ƶ�����
	 * @param today ���������
	 * @param samplePath ģ���·��
	 */
	static void getDataForDay(String today, String samplePath) {

		String myDate = "2015-" + today;
		String myFilePath = "D:/��������/����" + today + ".xls";
		
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
	 * ����ҳ�ϵ�����ȡ������ģ��
	 * @param date ����
	 * @param from �汾��
	 * @param wb Ŀ��洢�ļ�
	 * @param sheet Ŀ�����
	 * @param c ��ĵڼ����У���Ҫд0ΪԴ����ҳ��
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
