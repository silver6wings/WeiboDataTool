package dataTools;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;

import org.htmlparser.NodeFilter;
import org.htmlparser.Parser;
import org.htmlparser.filters.NodeClassFilter;
import org.htmlparser.tags.TableColumn;
import org.htmlparser.tags.TableHeader;
import org.htmlparser.tags.TableRow;
import org.htmlparser.tags.TableTag;
import org.htmlparser.util.NodeList;

public class ThTagDemo {
	
	public static String trim(String str){
		if(str.equals("-")) return "0";
		
		str = str.replace(",", "");
		str = str.replace("%", "");
		return str;
	}
	
	public static void main(String[] args) throws Exception {
		
		/*
		String html=""+		
		"<html>" +
		"<head>" +
		"</head>" +
		"<body>" +
		"<table>" +
		"<tr><th>th_Item</th></tr>" +
		"<tr><td>td_Item</td></tr>" +
		"</table>" +
		"</body>" +
		"</html>";		
		Parser parser = new Parser();
		parser.setInputHTML(html);
		parser.setEncoding("gbk");
		*/
		
		
		String strURL = "http://172.16.193.178/utils/incr_stats2?date=2014-10-19&from=1045595010";		
		Parser parser = new Parser((HttpURLConnection) (new URL(strURL)).openConnection());
		parser.setEncoding("utf-8");
		
		NodeFilter filter = new NodeClassFilter(TableTag.class);
		NodeList nodeList = parser.parse(filter);
		
		for(int i = 0; i < nodeList.size(); ++i){
			if(nodeList.elementAt(i) instanceof TableTag){
				TableTag tag = (TableTag) nodeList.elementAt(i);
				TableRow[] rows = tag.getRows();

				for (int j = 0; j < rows.length; ++j) {
					TableRow row = (TableRow) rows[j];
					// the reason to get headers is to parse <th> tag
					TableHeader[] headers = row.getHeaders();
					for (int k = 0; k < headers.length; ++k) {
//						 System.out.println("tag±êÇ©Îª£º" + headers[k].getTagName());
						 System.out.println(headers[k].getStringText());
					}
					
					TableColumn[] columns = row.getColumns();
					for (int k = 0; k < columns.length; ++k) {
						String info = columns[k].toPlainTextString().trim();
						System.out.println(Float.parseFloat(trim(info)));
					} // end for k
					
				} // end for j
			}
		}
		
	}
	
	public static String captureHtml(String strURL) throws Exception 
	{
		URL url = new URL(strURL);
		HttpURLConnection httpConn = (HttpURLConnection) url.openConnection();
		InputStreamReader input = new InputStreamReader(httpConn.getInputStream(), "utf-8");
		BufferedReader bufReader = new BufferedReader(input);
		
		StringBuilder contentBuf = new StringBuilder();
		String line = "";
		while ((line = bufReader.readLine()) != null) contentBuf.append(line);
		
		String buf = contentBuf.toString();
		return buf;
	}
	
	
}
