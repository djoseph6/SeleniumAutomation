package input.Excel;

import java.io.FileInputStream;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.log4testng.Logger;


public class RepaymentDetailsSheet {
//	Members	
	private static final String sheetName = "Repayment Details";
	private static FileInputStream excelFile;	
	private static final Logger logger = Logger.getLogger(RepaymentDetailsSheet.class);
	private static XSSFWorkbook workBook;
	public static XSSFSheet repayDetails;
	
	public static void prepareExcelFileAndSheet() throws IOException {
		String excelFilePath =".\\DataFiles\\Land Acquisition.xlsx";
			logger.info("Inputing excel file location into fileinputstream");
		excelFile = new FileInputStream(excelFilePath);
		workBook = new XSSFWorkbook(excelFile);
		repayDetails = workBook.getSheet(sheetName);
	}
	
	public static void interateThroughExcelSheet(XSSFSheet sheet){
		int rows = sheet.getLastRowNum();
		int colms = sheet.getRow(0).getLastCellNum();
		
		forloop1: for(int r=2; r<rows; r++){
			XSSFRow row = sheet.getRow(r);
			if(!row.equals(null)) {
				forloop2: for(int c=0; c<colms; c++) {
					XSSFCell cell = row.getCell(c);
					
					if(!cell.equals(null)) {
						switch(cell.getCellType()) {
							case STRING: System.out.println(cell.getStringCellValue()); break;
							case NUMERIC: System.out.println(cell.getNumericCellValue()); break;
							case BOOLEAN: System.out.println(cell.getBooleanCellValue()); break;
						}
					} else {
						continue forloop2;
					}
				}
					
			}else {
				continue forloop1;
			}
		}
			
	}
}
	


