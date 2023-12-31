package PJInvoice_Test;

import java.io.FileInputStream;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;
import Base.Browser;
import InvoiceS.PJInvoice1;
import InvoiceS.PJInvoice2;
import Utility.CommonFile;

public class PJTest extends Browser {

	PJInvoice1 pj1;
	PJInvoice2 pj2;

	@BeforeMethod
	public void setup() throws Exception {

		initilization();
		pj1 = new PJInvoice1();
		pj2 = new PJInvoice2();
		pj1.verifyLoginApp();
		Thread.sleep(2000);

		pj1.verifyIFFBtn();
		Thread.sleep(2000);
		pj1.verifyFinanceBtn();
		Thread.sleep(2000);
		pj1.verifyFreightJournalBtn();
		Thread.sleep(2000);
		pj2.verifyNewBtn();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

	}

	@Test(enabled = true)
	public void data() throws Exception {
		FileInputStream file1 = new FileInputStream(
				"C:\\Users\\Admin\\eclipse-workspace\\Freight_Journal\\Test_data\\Freight_Journal.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(file1);
		XSSFSheet sheet = workbook.getSheet("NewFreightJournal");
		int rowcount = sheet.getLastRowNum();
		int row = rowcount - 6;
		int colcount = sheet.getRow(7).getLastCellNum();
		System.out.println("rowcount in test:" + row + " colcount in test:" + colcount);

		for (int exec = 1; exec <= row; exec++) {
			Thread.sleep(2000);

			pj2.newFreightJournal(exec);
			pj2.basicDetail(exec);
			pj2.addDetails(exec);
			pj2.save();
			System.out.println("*** FREIGHT JOURNAL INVOICE DONE : "+exec+" ***");

		}

	}

	@AfterMethod

	public void exit(ITestResult b) throws Throwable
	{
		if(ITestResult.FAILURE == b.getStatus())
		{	
			CommonFile.captureScreenshotFaildTC(driver,b.getName());
		}
		Thread.sleep(2500);
	//	driver.quit();
	
	}
}
