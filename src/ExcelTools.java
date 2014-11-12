import java.io.FileOutputStream;
import java.text.DecimalFormat;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class ExcelTools {

	// 向表里写入数字
	public static void writeNumTo(double d, HSSFWorkbook workbook, String sheet, int row, int column) {
		try {
			if(null == workbook.getSheet(sheet)) workbook.createSheet(sheet);
			HSSFSheet tSheet = workbook.getSheet(sheet);
		
			if(null == tSheet.getRow(row)) tSheet.createRow(row);
			HSSFRow tRow = tSheet.getRow(row);
			
			if (null == tRow.getCell(column)) tRow.createCell(column);	
			HSSFCell tCell = tRow.getCell(column);	
			
			tCell.setCellValue(d);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	// 向表里写入字符串
	public static void writeStrTo(String s, HSSFWorkbook workbook, String sheet, int row, int column) {
		try {
			if(null == workbook.getSheet(sheet)) workbook.createSheet(sheet);
			HSSFSheet tSheet = workbook.getSheet(sheet);
		
			if(null == tSheet.getRow(row)) tSheet.createRow(row);
			HSSFRow tRow = tSheet.getRow(row);
			
			if (null == tRow.getCell(column)) tRow.createCell(column);	
			HSSFCell tCell = tRow.getCell(column);	
			
			tCell.setCellValue(s);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 输出到文件
	public static void workbookOut(HSSFWorkbook workbook, String file) {
		try {
			FileOutputStream fOut = new FileOutputStream(file);
			workbook.write(fOut);
			fOut.flush();
			fOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void copySheetPart(HSSFSheet sourceSheet, int sr, int sc, 
			HSSFSheet targetSheet, int tr, int tc,
			FormulaEvaluator evaluator){
		copySheetPart(sourceSheet, sr, sc, targetSheet, tr, tc, 1, 1, evaluator);
	}
	
	// Sheet区域矩阵复制
	public static void copySheetPart(HSSFSheet sourceSheet, int sr, int sc, 
			HSSFSheet targetSheet, int tr, int tc, int rows, int columns, 
			FormulaEvaluator evaluator){
		for (int i = 0; i < rows; i++) {
			if (null == sourceSheet.getRow(sr + i)) continue;
			
			for (int j = 0; j < columns; j++) {
				
				HSSFCell sCell = sourceSheet.getRow(sr + i).getCell(sc + j);
				if(null == sCell) continue;			
				
				HSSFRow tRow;
				if (null == targetSheet.getRow(tr + i)) tRow = targetSheet.createRow(tr + i);
				else tRow = targetSheet.getRow(tr + i);
					
				HSSFCell tCell;
				if (null == tRow.getCell(tc+j)) tCell = tRow.createCell(tc+j);
				else tCell = tRow.getCell(tc+j);
								
				switch (sCell.getCellType()) {			
				case HSSFCell.CELL_TYPE_NUMERIC:
					double x = sCell.getNumericCellValue();
					tCell.setCellValue(x);		
					break;
				case HSSFCell.CELL_TYPE_STRING:
					tCell.setCellValue(sCell.getStringCellValue());
					break;
				case HSSFCell.CELL_TYPE_FORMULA:
					evaluator.evaluateFormulaCell(sCell);  
                    tCell.setCellValue(sCell.getNumericCellValue());
					break;
				default:										
					tCell.setCellValue("Unknown Type");
				}				
			}
		}
	}
	
	// 获得一个Cell矩阵并输出
	public static void getCellsArray(HSSFSheet tSheet, int r, int c, int h, int w) {

		if (r < 0 || c < 0 || h < 0 || w < 0) return;

		for (int i = r; i < r+h; i++) {
			HSSFRow tRow = tSheet.getRow(i);
			for (int j = c; j < c+w; j++) {
				HSSFCell tCell = tRow.getCell(j);
				System.out.print(getCellContent(tCell) + "\t\t");
			}
			System.out.println();
		}
	}

	// 将Cell内容转化为字符串
	public static String getCellContent(HSSFCell tCell) {

		if (tCell == null)
			return "null";

		switch (tCell.getCellType()) {

		case 0:// Numeric
			DecimalFormat df = new DecimalFormat("#");
			return df.format(tCell.getNumericCellValue());

		case 1:// String
			return tCell.getStringCellValue();

		default:
			return "UnknownType";
		}
	}
}