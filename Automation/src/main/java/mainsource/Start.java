package mainsource;

import java.io.IOException;

import input.Excel.RepaymentDetailsSheet;

public class Start {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		RepaymentDetailsSheet.prepareExcelFileAndSheet();
		RepaymentDetailsSheet.interateThroughExcelSheet(RepaymentDetailsSheet.repayDetails);
	}

}
