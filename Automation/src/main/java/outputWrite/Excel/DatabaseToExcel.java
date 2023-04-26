package outputWrite.Excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;


import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.log4testng.Logger;

import util.Classes.*;
public class DatabaseToExcel {
	
	private static final Logger log = Logger.getLogger(DatabaseToExcel.class);
	private static ConnectionUtil connUtil;
	private static Connection conn;
	private static ResultSet rs;
	private static XSSFWorkbook workbook;
	private static XSSFSheet worksheet;
	private static FileOutputStream outStream;
	private static String[] databaseColmTitles = new String[]
			{"LOCATION_ID", "STREET_ADDRESS", "POSTAL_CODE", "CITY", "STATE_PROVINCE","COUNTRY_ID"};
	
	public static void prepareConnection() {
		connUtil = ConnectionUtil.getConnectionUtil();
		conn = connUtil.getConnection();
	}
	
	public static ResultSet getResultSetForQuery(String databaseQuery) throws SQLException {
		Statement stmt = conn.createStatement();
		rs = stmt.executeQuery(databaseQuery);
		
		return rs;
	}
	
	public static void prepareWorkbookWithColumnNames() {
		 workbook = new XSSFWorkbook();
		 worksheet = workbook.createSheet("Location Data");
		 XSSFRow row = worksheet.createRow(0);
		 	for(int a=0; a<databaseColmTitles.length; a++) {
		 		row.createCell(a).setCellValue(databaseColmTitles[a]);
		 	}
		 
	}
	
	public static void interateDatabaseResultSetToExcel() throws SQLException, IOException {
		if(rs!=null) {
			int newRowCounter=0;
			while(rs.next()) {
				newRowCounter++;
				
				double locId = rs.getDouble(databaseColmTitles[0]);
				String strAdd = rs.getString(databaseColmTitles[1]);
				String pstCode = rs.getString(databaseColmTitles[2]);
				String cty = rs.getString(databaseColmTitles[3]);
				String stateProv = rs.getString(databaseColmTitles[4]);
				String countryId = rs.getString(databaseColmTitles[5]);
				
				XSSFRow row = worksheet.createRow(newRowCounter);
				forloop: for(int cellNum=0; cellNum<=databaseColmTitles.length; cellNum++) {
					switch(cellNum) {
						case 0: row.createCell(cellNum).setCellValue(locId); break;
						case 1: row.createCell(cellNum).setCellValue(strAdd); break;
						case 2: row.createCell(cellNum).setCellValue(pstCode); break;
						case 3: row.createCell(cellNum).setCellValue(cty); break;
						case 4: row.createCell(cellNum).setCellValue(stateProv); break;
						case 5: row.createCell(cellNum).setCellValue(countryId); break;
						default: continue forloop;
					}
				}
			}
				createAndWriteWorkbookFile();
		}
	}
	
	private static void createAndWriteWorkbookFile() throws IOException {
		String filePath = ".\\DataFiles\\locations.xlsx";	//set file path where file will be created
		
		log.info("Attempting to create and write Workbook file");		//log attempt to create and write workbook file
		try {
			outStream = new FileOutputStream(filePath);	//inputing file location to output stream
			workbook.write(outStream);	//try to create workbook file at filepath location
		}catch(Exception e) {
			e.getStackTrace();	//print exception if attempt fails
		}finally {
			outStream.close();
		}
		
	}
}
