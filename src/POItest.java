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
			// �����µ�Excel ������
			HSSFWorkbook workbook = new HSSFWorkbook();
			// ��Excel�������н�һ����������Ϊȱʡֵ
			// ��Ҫ�½�һ��Ϊ"Ч��ָ��"�Ĺ����������Ϊ��
			
			HSSFSheet sheet = workbook.createSheet("asdfasdf");

			// ������0��λ�ô����У���˵��У�
			HSSFRow row = sheet.createRow(0);
			// ������0��λ�ô�����Ԫ�����϶ˣ�
			HSSFCell cell = row.createCell(0);
			
			// ���嵥Ԫ��Ϊ�ַ�������
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			// �ڵ�Ԫ��������һЩ����
			cell.setCellValue("sweater");
			
			String outputFile = "d:/feed-daily-report-v2-201407302.xls";
			// �½�һ����ļ���
			FileOutputStream fOut = new FileOutputStream(outputFile);
			// ����Ӧ��Excel ����������
			workbook.write(fOut);
			fOut.flush();
			// �����������ر��ļ�
			fOut.close();
			System.out.println("�ļ�����...");

		} catch (Exception e) {
			System.out.println("������ xlCreate() : " + e);
		}
	}
	

	/**
	 * 
	 * ��ȡexcel����������С���ȡ������Ϣ�����ж����Ƿ����ֻ����룬������ȷ���ֻ����������ʾ
	 * 
	 * 
	 * ע�⣺ 1.sheet�� ��0��ʼ����workbook.getNumberOfSheets()-1���� 
	 * 2.row��
	 * ��0��ʼ(getFirstRowNum)����getLastRowNum���� 3.cell��
	 * ��0��ʼ(getFirstCellNum)����getLastCellNum����, ��������Ŀ��֪ʲôԭ������ʾ�ĳ��Ȳ�ͬ�����ܻ�ƫ��
	 * 
	 */
	public void readExcel() {

		try {
			String fileToBeRead = "d:/feed-daily-report-v2-20140730.xls";
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileToBeRead));
			
			System.out.println(">>>>"+workbook.getNumberOfSheets());
			
			for (int numSheets = 0; numSheets <= 0; numSheets++) {
				if (null != workbook.getSheetAt(numSheets)) {
					HSSFSheet aSheet = workbook.getSheetAt(numSheets);// ���һ��sheet			
					
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
										// System.out.println("��ʽ���Բ���");//������ʽ������
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