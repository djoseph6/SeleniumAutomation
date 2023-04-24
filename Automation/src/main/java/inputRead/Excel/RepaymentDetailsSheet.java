package inputRead.Excel;

import java.io.FileInputStream;

import java.io.IOException;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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
		String excelFilePath =".\\DataFiles\\Land Acquisition.xlsx";		//Set file XL file location to string value
			logger.info("Inputing excel file location into fileinputstream");		//Log attempt to fileinputstream with string location
		excelFile = new FileInputStream(excelFilePath); 		//Open connection in FileInputStream object with XL file 
		workBook = new XSSFWorkbook(excelFile); 		//Create workbook object from XL file
//		workBook = (XSSFWorkbook) WorkbookFactory.create(excelFile, "password for excel");		//used if there is a password for excel file
		repayDetails = workBook.getSheet(sheetName);		//Find worksheet with title "Repayment Details" from workbook and create worksheet object from it
		excelFile.close();		//Close fileinputstream object
	}
	
	public static void interateThroughExcelSheet1(XSSFSheet sheet){
		int rows = sheet.getLastRowNum();		//get how many rows are in the sheet
		int colms = sheet.getRow(0).getLastCellNum();		//get how many columns are in the sheet
		
		forloop1: for(int r=0; r<rows; r++){		//Outer forloop to interate through sheet rows
			XSSFRow row = sheet.getRow(r);		//For each worksheet row, create row object from it
			if(!row.equals(null)) {			//if the row is not null
				forloop2: for(int c=0; c<colms; c++) {			//Inner forloop to iterate through sheet columns for each row AKA the cells
					XSSFCell cell = row.getCell(c);			//Get value from column index in the row AKA get value from the cell, create and assign it to cell object
					
					if(!cell.equals(null)) {		//if the cell is not null
						switch(cell.getCellType()) {		//retrieve cell type and read it
							case STRING: System.out.println(cell.getStringCellValue()); break;		//If string, then print the string value
							case NUMERIC: System.out.println(cell.getNumericCellValue()); break;	//If number, then print the numeric value	
							case BOOLEAN: System.out.println(cell.getBooleanCellValue()); break;	//If boolean, then print the boolean value
							case FORMULA: 		//If formulated value,
								try{
									System.out.println(cell.getNumericCellValue()); break;		//try printing as numeric value
								} catch(Exception e){		//If not numeric value, catch exception
									e.printStackTrace();
										try {
											System.out.println(cell.getStringCellValue()); break;		//try printing as String value
										}catch(Exception f) {		//If not string value, catch exception
											f.printStackTrace();
												try {
													System.out.println(cell.getBooleanCellValue()); break;		//try printing as boolean value
												}catch(Exception g) {
													g.printStackTrace();
													logger.debug("Formulated Cell value is neither numeric, string or boolean");
												}
											
										}
									
								}
						}
					} else {		//if cell is null, 
						continue forloop2;		//skip this cell
					}
				}
					
			}else {			//if row is null,
				continue forloop1;			//skip this row
			}
		}
			
	}
	
	public static void interateThroughExcelSheet2(XSSFSheet sheet) {
		if(sheet!= null) {		//if passed in sheet is not null
			Iterator<Row> iterator = sheet.rowIterator();	//Create row iterator object for the worksheet
			while(iterator.hasNext()) {		//If there is a next row in xl sheet
				XSSFRow row = (XSSFRow) iterator.next();	//create a row object from the next row
				Iterator<Cell> cellIterator = row.cellIterator();	//Create cell iterator object for the row
				while(cellIterator.hasNext()) {		//If there is a next cell in row
					XSSFCell cell = (XSSFCell) cellIterator.next();		//create a cell object from the next cell
					
					if(!cell.equals(null)) {		//if cell is not null
						
						switch(cell.getCellType()) {
							case STRING: System.out.println(cell.getStringCellValue()); break;
							case NUMERIC: System.out.println(cell.getNumericCellValue()); break;
							case BOOLEAN: System.out.println(cell.getBooleanCellValue()); break;
							case FORMULA: try{
								System.out.println(cell.getNumericCellValue()); break;		//try printing as numeric value
							} catch(Exception e){		//If not numeric value, catch exception
								e.printStackTrace();
									try {
										System.out.println(cell.getStringCellValue()); break;		//try printing as String value
									}catch(Exception f) {		//If not string value, catch exception
										f.printStackTrace();
											try {
												System.out.println(cell.getBooleanCellValue()); break;		//try printing as boolean value
											}catch(Exception g) {
												g.printStackTrace();
												logger.debug("Formulated Cell value is neither numeric, string or boolean");
											}
										
									}
								
							}
						}
					}
					
				}
				System.out.println();
			}
			
		}
		
	}
}
	


