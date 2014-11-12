import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DecimalFormat;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;

public class POItest {
	
	public void CreateExcel() {
		try {
			// 创建新的Excel 工作簿
			HSSFWorkbook workbook = new HSSFWorkbook();
			// 在Excel工作簿中建一工作表，其名为缺省值
			// 如要新建一名为"效益指标"的工作表，其语句为：
			
			HSSFSheet sheet = workbook.createSheet("asdfasdf");

			// 在索引0的位置创建行（最顶端的行）
			HSSFRow row = sheet.createRow(0);
			// 在索引0的位置创建单元格（左上端）
			HSSFCell cell = row.createCell(0);
			
			// 定义单元格为字符串类型
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			// 在单元格中输入一些内容
			cell.setCellValue("sweater");
			
			String outputFile = "d:/feed-daily-report-v2-201407302.xls";
			// 新建一输出文件流
			FileOutputStream fOut = new FileOutputStream(outputFile);
			// 把相应的Excel 工作簿存盘
			workbook.write(fOut);
			fOut.flush();
			// 操作结束，关闭文件
			fOut.close();
			System.out.println("文件生成...");

		} catch (Exception e) {
			System.out.println("已运行 xlCreate() : " + e);
		}
	}
	

	/**
	 * 
	 * 读取excel，遍历各个小格获取其中信息，并判断其是否是手机号码，并对正确的手机号码进行显示
	 * 
	 * 
	 * 注意： 1.sheet， 以0开始，以workbook.getNumberOfSheets()-1结束 
	 * 2.row，
	 * 以0开始(getFirstRowNum)，以getLastRowNum结束 3.cell，
	 * 以0开始(getFirstCellNum)，以getLastCellNum结束, 结束的数目不知什么原因与显示的长度不同，可能会偏长
	 * 
	 */
	public void readExcel() {

		try {
			String fileToBeRead = "d:/feed-daily-report-v2-20140730.xls";
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileToBeRead));
			
			System.out.println(">>>>"+workbook.getNumberOfSheets());
			
			for (int numSheets = 0; numSheets <= 0; numSheets++) {
				if (null != workbook.getSheetAt(numSheets)) {
					HSSFSheet aSheet = workbook.getSheetAt(numSheets);// 获得一个sheet			
					
					// System.out.println("+++getFirstRowNum+++" +
					// aSheet.getFirstRowNum());//
					// System.out.println("+++getLastRowNum+++" +
					// aSheet.getLastRowNum());

					for (int rowNumOfSheet = 0; rowNumOfSheet <= aSheet.getLastRowNum(); rowNumOfSheet++) {
						if (null != aSheet.getRow(rowNumOfSheet)) {
							HSSFRow aRow = aSheet.getRow(rowNumOfSheet);

							// System.out.println(">>>getFirstCellNum<<<"+
							// aRow.getFirstCellNum());
							// System.out.println(">>>getLastCellNum<<<"+
							// aRow.getLastCellNum());
							for (int cellNumOfRow = 0; cellNumOfRow <= aRow.getLastCellNum(); cellNumOfRow++) {

								if (null != aRow.getCell(cellNumOfRow)) {
									HSSFCell aCell = aRow.getCell(cellNumOfRow);

									
									int cellType = aCell.getCellType();
									// System.out.println(cellType);
									switch (cellType) {
									case 0:// Numeric
										DecimalFormat df = new DecimalFormat("#");
										String strCell = df.format(aCell
												.getNumericCellValue());

										System.out.println(strCell);

										break;
									case 1:// String
										strCell = aCell.getStringCellValue();

										System.out.println(strCell);

										break;
									default:
										// System.out.println("格式不对不读");//其它格式的数据
									}
								} else {
									System.out.println("null");
								}
							}
						}
					}
				}
			}
		} catch (Exception e) {
			System.out.println("ReadExcelError" + e);
		}
	}

	public void Test(int row, int colmun){


	}
	
	public static void main(String[] args) {
		
		POItest poi = new POItest();
		// poi.CreateExcel();
//		poi.readExcel();
		poi.CreateExcel();
		
		System.out.println("1324");
	}
}