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

	static String today = "11-23";
	static String myDate = "2014-" + today;
	static String myFilePath = "D:/数据增量/增量数据" + today + ".xls";
	static HSSFWorkbook myWorkbook;

	public static void main(String[] args) {

		System.out.println("Start:");
		try {
			myWorkbook = new HSSFWorkbook(new FileInputStream(myFilePath));
			getData(myDate, "1045595010", "Android", 0);
			getData(myDate, "1046095010", "Android", 1);
			getData(myDate, "1046195010", "Android", 2);
			getData(myDate, "1046295010", "Android", 3);
			
			getData(myDate, "1046093010", "iPhone", 0);
			getData(myDate, "1046193010", "iPhone", 1);
			getData(myDate, "1046593010", "iPhone", 2);
			
			for (int i = 0; i < 3; i++) 
				myWorkbook.getSheetAt(i).setForceFormulaRecalculation(true);
						
			System.out.println("Outing...");
			ExcelTools.workbookOut(myWorkbook, myFilePath);
			System.out.println("Done!");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	static void getData(String date, String from, String sheet, int c) throws Exception {
		int w = 8;
		String strURL = "http://172.16.193.178/utils/incr_stats2?date=" + date + "&from=" + from;

		Parser parser = new Parser((HttpURLConnection) (new URL(strURL)).openConnection());
		parser.setEncoding("utf-8");

		NodeFilter filter = new NodeClassFilter(TableTag.class);
		NodeList nodeList = parser.parse(filter);
		
		ExcelTools.writeStrTo(date, myWorkbook, sheet, 0, w * c);
		ExcelTools.writeStrTo(from, myWorkbook, sheet, 1, w * c);
		
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

							ExcelTools.writeNumTo(parseContent(info), myWorkbook, sheet, 5 + j, w * c + k + 1);
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

							ExcelTools.writeNumTo(parseContent(info), myWorkbook, sheet, 10, w * c + k);
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

							ExcelTools.writeNumTo(parseContent(info), myWorkbook, sheet, 12 + j, w * c + k + 1);
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
							ExcelTools.writeNumTo(parseContent(info), myWorkbook, sheet, 19 + j, w * c + k + 1);
						}
					}

					int j = rows.length-1;
					TableRow row = (TableRow) rows[j];
					TableColumn[] columns = row.getColumns();
					for (int k = 0; k < columns.length; ++k) {
						String info = columns[k].toPlainTextString().trim();
						System.out.println(info);
						ExcelTools.writeNumTo(parseContent(info), myWorkbook, sheet, 70, w * c + k + 1);
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
