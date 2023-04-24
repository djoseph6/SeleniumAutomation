package outputWrite.Excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.log4testng.Logger;

public class WriteToExcel {
	
	private static final String sheetName = "Data Output";
	private static XSSFWorkbook workBook; //creating a empty excel workbook
	private static XSSFSheet sheet;
	private static FileInputStream excelFile;
	private static FileOutputStream outStream;		//initialize FOS object
	private static final Logger log = Logger.getLogger(WriteToExcel.class);
	private static Object empData[][] = {
							{"EmpId","Name", "Job"},  //[1]
							{101, "David", "Lawyer"},  //[2]
							{102,"Mitch", "Teacher"},  //[3]
							{103, "Toni", "Therapist"}  //[4]
						};
	
	//Supporting method for addDataToExcelSheetAndCreateFile1 and 2. Returns ArrayList
	public static ArrayList<Object[]> createAndReturnDataArrayList() {
		 ArrayList<Object[]> empData1 = new ArrayList<Object[]>();
			empData1.add(new Object[] {"EmpId","Name", "Job"});
			empData1.add(new Object[] {101, "David", "Lawyer"});
			empData1.add(new Object[] {102,"Mitch", "Teacher"});
			empData1.add(new Object[] {103, "Toni", "Therapist"});
			
			return empData1;
	 }
	//Creates workbook and worksheet to work with.
	public static void createExcelFileAndSheet() throws IOException {
		workBook = new XSSFWorkbook();
			log.info("Creating new workbook");		//Log attempt to create new workbook
		sheet = workBook.createSheet(sheetName); //creating sheet in xl workbook called "Data Output"
	}
	
	
	
	//Write 2D Array to XL sheet
	public static void addDataToExcelSheetAndCreateFile1(Object[][] inputData) throws IOException {
		int rows = inputData.length;	//4
		int colms = inputData[0].length;  //3
		
		for(int r=0;r<rows;r++) {	//Using for loop
			XSSFRow row = sheet.createRow(r);  //create a row within the sheet
			for(int c=0;c<colms;c++) {
				XSSFCell cell = row.createCell(c);	//create a cell within the row
				Object value = inputData[r][c];	//assign array value to holding object
				
				if(value instanceof String) {
					cell.setCellValue((String)value);	//write string value in cell
				} else if(value instanceof Integer) {
					cell.setCellValue((Integer)value);	//write integer value in cell
				} else if(value instanceof Boolean) {
					cell.setCellValue((Boolean)value);	//write boolean value in cell
				}
			}
		}
		createAndWriteWorkbookFile();
	}
	
	//Write ArrayList<Objec[]> to XL sheet
	public static void addDataToExcelSheetAndCreateFile2() throws IOException {
		int rowCount =0;
		for(Object[] rows: empData) {	//seperate nested array by rows. 4 rows in all.
			XSSFRow row = sheet.createRow(rowCount++);	//create a new row within sheet at rowCount index
			int colmCount=0;
			for(Object value: rows) {	//iterate through values of array. 3 in all.
				XSSFCell cell = row.createCell(colmCount++);	//create a cell within row at colmCount index
				
				if(value instanceof String) {
					cell.setCellValue((String)value);	//write string value in cell
				} else if(value instanceof Integer) {
					cell.setCellValue((Integer)value);	//write integer value in cell
				} else if(value instanceof Boolean) {
					cell.setCellValue((Boolean)value);	//write boolean value in cell
				}
			}
		}
		createAndWriteWorkbookFile();
	}
	
	//Write ArrayList<Objec[]> to XL sheet
	public static void addDataToExcelSheetAndCreateFile2(ArrayList<Object[]> empData) throws IOException {
		int rowCount =0;
		for(Object[] rows: empData) {	//seperate nested array by rows. 4 rows in all.
			XSSFRow row = sheet.createRow(rowCount++);	//create a new row within sheet at rowCount index
			int colmCount=0;
			for(Object value: rows) {	//iterate through values of array. 3 in all.
				XSSFCell cell = row.createCell(colmCount++);	//create a cell within row at colmCount index
				
				if(value instanceof String) {
					cell.setCellValue((String)value);	//write string value in cell
				} else if(value instanceof Integer) {
					cell.setCellValue((Integer)value);	//write integer value in cell
				} else if(value instanceof Boolean) {
					cell.setCellValue((Boolean)value);	//write boolean value in cell
				}
			}
		}
		createAndWriteWorkbookFile();
	}
	
	//Adds boolean value if employee id is greater than 101 to each row in sheet
	public static void addFormulaCellToExcelSheet() throws IOException {
		if(sheet!= null) {		//if passed in sheet is not null
			int rows = sheet.getLastRowNum();		//get how many rows are in the sheet
			
			forloop1: for(int r=1; r<rows+1; r++){		//Outer forloop to interate through sheet rows
				XSSFRow row = sheet.getRow(r);		//For each worksheet row, create row object from it
				int lastCellNum = row.getLastCellNum();		//assign last row cell index to variable
				row.createCell(lastCellNum, CellType.FORMULA);		//create a new cell at last cell index and assign cell type as formula
				row.getCell(lastCellNum).setCellFormula("A"+(r+1)+">101");		//set formula to compare column A(empid) to 101
			}
			createAndWriteWorkbookFile();
		}else {
			log.debug("Sheet is null when adding formula to XL sheet");
		}
	}
	
	//Supporting methods for addDataToExcelSheetAndCreateFile1 and 2.
	private static void createAndWriteWorkbookFile() throws IOException {
		String filePath = ".\\DataFiles\\employees.xlsx";	//set file path where file will be created
		
		log.info("Attempting to create and write Workbook file");		//log attempt to create and write workbook file
		try {
			outStream = new FileOutputStream(filePath);	//inputing file location to output stream
			workBook.write(outStream);	//try to create workbook file at filepath location
		}catch(Exception e) {
			e.getStackTrace();	//print exception if attempt fails
		}finally {
			outStream.close();	//close fileoutputstream object
		}
		
	}
	
	
	
}
