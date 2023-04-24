package mainsource;

import java.io.IOException;
import java.util.ArrayList;

import inputRead.Excel.RepaymentDetailsSheet;
import outputWrite.Excel.WriteToExcel;

public class Start {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		//RepaymentDetailsSheet.prepareExcelFileAndSheet();
		//RepaymentDetailsSheet.interateThroughExcelSheet2(RepaymentDetailsSheet.repayDetails);
		
		WriteToExcel.createExcelFileAndSheet();
		ArrayList<Object[]> dataObj = WriteToExcel.createAndReturnDataArrayList();
		WriteToExcel.addDataToExcelSheetAndCreateFile2(dataObj);
		WriteToExcel.addFormulaCellToExcelSheet();
	}

}
