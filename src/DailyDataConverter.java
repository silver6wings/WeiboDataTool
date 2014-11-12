import java.io.FileInputStream;
import java.io.FileOutputStream;


//import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class DailyDataConverter {

	public static String targetFilePath = "d:/数据/趋势daily下发量汇总20141111.xls";
	public FormulaEvaluator e;
    
	
	public static void main(String[] args) {
		
		DailyDataConverter ddc = new DailyDataConverter();
		System.out.println(ddc.doConvert("d:/数据/feed-daily-report-v2-20141111.xls", targetFilePath, 27));
//		System.out.println(ddc.doConvert("d:/数据/feed-daily-report-v2-20141107.xls", targetFilePath, 27));
//		System.out.println(ddc.doConvert("d:/数据/feed-daily-report-v2-20141108.xls", targetFilePath, 27));
//		System.out.println(ddc.doConvert("d:/数据/feed-daily-report-v2-20141109.xls", targetFilePath, 27));
//		System.out.println(ddc.doConvert("d:/数据/feed-daily-report-v2-20141110.xls", targetFilePath, 27));
//		System.out.println(ddc.doConvert("d:/数据/feed-daily-report-v2-20141103.xls", targetFilePath, 27));
//		System.out.println(ddc.doConvert("d:/数据/feed-daily-report-v2-20141104.xls", targetFilePath, 27));
	}

	// 进行转换
	public int doConvert(String sourceFile, String targetFile, int datePos) {
	
		try {
			HSSFWorkbook targetWorkbook = new HSSFWorkbook(new FileInputStream(targetFile));
			e = targetWorkbook.getCreationHelper().createFormulaEvaluator();
			 
			// 整理信息
			HSSFSheet sheetTarget = targetWorkbook.getSheet("Data-edit");
			HSSFWorkbook sourceWorkbook = new HSSFWorkbook(new FileInputStream(sourceFile));
			
			HSSFSheet sheetAll = sourceWorkbook.getSheet("All");
			ExcelTools.copySheetPart(sheetAll, 1, 1, sheetTarget, 3, 2, 5, 5, e);			
			
			HSSFSheet sheet10 = sourceWorkbook.getSheet("source_10_ADs");
			ExcelTools.copySheetPart(sheet10, 1, 1, sheetTarget, 4, 9, 5, 3, e);	
			
			HSSFSheet sheet130 = sourceWorkbook.getSheet("source_130");
			ExcelTools.copySheetPart(sheet130, 1, 1, sheetTarget, 4, 14, 5, 3, e);	
						
			HSSFSheet sheet131 = sourceWorkbook.getSheet("source_131");
			ExcelTools.copySheetPart(sheet131, 1, 1, sheetTarget, 13, 14, 5, 3, e);	
			
			HSSFSheet sheet20 = sourceWorkbook.getSheet("source_20_Trends");
			ExcelTools.copySheetPart(sheet20, 1, 1, sheetTarget, 13, 9, 5, 3, e);	
			
			HSSFSheet sheet122 = sourceWorkbook.getSheet("source_122");
			ExcelTools.copySheetPart(sheet122, 1, 1, sheetTarget, 22, 9, 5, 3, e);
			
			sheetTarget.setForceFormulaRecalculation(true);
						
			// 聚合信息
			HSSFSheet sheetFinal = targetWorkbook.getSheet("Data-new");
			int lastRow = sheetFinal.getLastRowNum();
			
			ExcelTools.copySheetPart(	sheetTarget, 30, 2,	sheetFinal, lastRow+1, 2, 5, 19, e);
			System.out.println("Completed");
			
			// 添加日期
			for (int i = 1; i <= 5; i++) {
				String s = sourceFile.substring(datePos, datePos+4) + 
						"/" + sourceFile.substring(datePos+4, datePos+6) + 
						"/" + sourceFile.substring(datePos+6, datePos+8);
				System.out.println(s);
				
				HSSFCell tCell = sheetFinal.getRow(lastRow+i).getCell(1);
				if (null == tCell) tCell = sheetFinal.getRow(lastRow+i).createCell(1);
				tCell.setCellValue(s);
			}
			
			// 加粗首行
			/*
			HSSFFont f = targetWorkbook.createFont();
			f.setFontHeightInPoints((short) 10);//字号 
			f.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			HSSFRow tRow =  sheetFinal.getRow(lastRow+1);
			
			for (int i = 1; i <= tRow.getLastCellNum(); i++) {
				HSSFCell tCell = tRow.getCell(i);
				if (null != tCell) tCell.getCellStyle().setFont(f);
			}
			*/
			
			// 输出			
			FileOutputStream fOut = new FileOutputStream(targetFile);
			targetWorkbook.write(fOut);
			fOut.flush();
			fOut.close();
			
		} catch (Exception e) {
			e.printStackTrace();
			return 1;
		}
		return 0;
	}

}
