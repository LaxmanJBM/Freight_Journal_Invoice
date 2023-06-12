package InvoiceS;

import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.Select;
import Base.Browser;


public class PJInvoice2 extends Browser{
	
	@FindBy(xpath="//img[@id='ctl00_btnNew']")private WebElement newBtn;
	@FindBy(xpath="//*[@id=\"divNewButton\"]/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/table/tbody//label")private List<WebElement> allRedioBtn1;
	@FindBy(xpath="//*[@id=\"divNewButton\"]/table/tbody/tr[2]/td[1]/table/tbody/tr[2]/td/table/tbody//label")private List<WebElement> allRedio;
	@FindBy(xpath="//*[@id=\"divNewButton\"]/table/tbody/tr[2]/td[1]/table/tbody/tr[3]/td/table/tbody//label")private List<WebElement> allRedioBtn3;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlOffice']")private WebElement office;	
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$imgVendor']")private WebElement partD;
	@FindBy(xpath="//input[@id='amp_common_search_lookup_textbox_control__0']")private WebElement partyText;
	@FindBy(xpath="//*[@id=\"amp_common_search_lookup_table_control_\"]/tbody//tr//td")private List<WebElement> allParty;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlPartyOU']")private WebElement partyOU;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlAccount']")private WebElement account;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlVouchor']")private WebElement voucherType;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$btnNewOk']")private WebElement okBtn;
	
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$ctl00']")private WebElement placeOfSupply;
	@FindBy(xpath="//*[@id=\"amp_common_search_lookup_table_control_\"]/tbody//tr//td")private List<WebElement> allPlaceOfSupply;
	@FindBy(xpath="//input[@id='amp_common_search_lookup_textbox_control__0']")private WebElement placeOfSupplyText;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlIncExp']")private WebElement serviceAccType;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlSvcAccount']")private WebElement partyAccType;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$btnou']")private WebElement partyouD;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlParty_OU']")private WebElement partuOU2;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$Btnouok']")private WebElement okBtn2;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlTaxType']")private WebElement taxType;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$txtDate']")private WebElement transDate;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$txtDueDate']")private WebElement dueDate;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$txtDocDate']")private WebElement docDate;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$txtARAPRef']")private WebElement arpRef;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlType']")private WebElement invoiceType;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$btnSaveDraft']")private WebElement saveAsDraft;
	
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$btnAdd_grid']")private WebElement addDetailsBtn;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$imgDetBook']")private WebElement bookD;
	@FindBy(xpath="//input[@id='amp_common_search_lookup_textbox_control__0']")private WebElement bookText;
	@FindBy(xpath="//*[@id=\"amp_common_search_lookup_table_control_\"]/tbody//tr//td")private List<WebElement> allBookRef;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$btn_tariff']")private WebElement tariffD;
	@FindBy(xpath="//input[@id='amp_common_search_lookup_textbox_control__0']")private WebElement tariffText;
	@FindBy(xpath="//*[@id=\"amp_common_search_lookup_table_control_\"]/tbody//tr//td")private List<WebElement> tariffList;
	@FindBy(xpath="//textarea[@name='ctl00$ContentPlaceHolder1$txt_desc']")private WebElement desc;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlUOM']")private WebElement uom;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$txtRate']")private WebElement rate;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$txtqty']")private WebElement qty;
	@FindBy(xpath="//select[@name='ctl00$ContentPlaceHolder1$ddlEstCur']")private WebElement currency;
	@FindBy(xpath="//textarea[@name='ctl00$ContentPlaceHolder1$txtEstRemarks']")private WebElement remark;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$btnAddGrdRow']")private WebElement addBtn;
	@FindBy(xpath="//textarea[@name='ctl00$ContentPlaceHolder1$txtRmks']")private WebElement finalRemark;
	
	@FindBy(xpath="//div[@class='fmBox ok']")private WebElement succMasg;
	@FindBy(xpath="//a[text()='Close']")private WebElement closeBtn;
	@FindBy(xpath="//img[@id='ctl00_btnSave']")private WebElement saveBtn;
	@FindBy(xpath="//img[@id='ctl00_btnCancel']")private WebElement undo;
	@FindBy(xpath="//img[@id='ctl00_btnNew']")private WebElement newBtnS;
	@FindBy(xpath="//div[@class='fmBox ok']")private WebElement jurnalSave;
	@FindBy(xpath="//a[text()='Close']")private WebElement closeBtnS;
	@FindBy(xpath="//input[@name='ctl00$ContentPlaceHolder1$imgclosediv']")private WebElement close;
	
/*	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;
	@FindBy(xpath="")private WebElement;*/
	
	public PJInvoice2() {
		PageFactory.initElements(driver, this);
	}
	
	public void verifyNewBtn() throws Exception {
		Set<String> window = driver.getWindowHandles();

		Iterator<String> it = window.iterator();

		String mainpage = driver.getWindowHandle();
		while (it.hasNext()) {
			String str = it.next();
			if (!mainpage.equals(str)) {
				driver.switchTo().window(str);
			}
		}

		newBtn.click();
	}
	
	
	public void newFreightJournal(int excel) throws Exception {
		
		  FileInputStream file5=new FileInputStream("C:\\Users\\Admin\\eclipse-workspace\\Freight_Journal\\Test_data\\Data1.xlsx");	
			
			
				XSSFWorkbook workbook=new XSSFWorkbook(file5);
				XSSFSheet sheet = workbook.getSheet("NewFreightJournal");
				int rowcount = sheet.getLastRowNum();
				
				int colcount = sheet.getRow(7).getLastCellNum();
				System.out.println("NewFreightJournal rowcount:"+rowcount+"NewFreightJournal colcount"+colcount);

				for(int i=7;i<=rowcount;i++)
				{
					XSSFRow celldata = sheet.getRow(i);	
					try {
					System.out.println("VALUE OF ID ="+ celldata.getCell(1).getNumericCellValue());
					int idNo = (int) celldata.getCell(1).getNumericCellValue();
					
					if(idNo == excel) {
						
//PJ All Redio Button
						try {
							
							String val = celldata.getCell(2).getStringCellValue();
							Thread.sleep(1000);
							for(int i1=0;i1<=allRedioBtn1.size();i1++) {
					//			System.out.println("List of First Redio Buttons ="+allRedioBtn1.get(i1).getText());
							if(allRedioBtn1.get(i1).getText().equalsIgnoreCase(val)) {
								allRedioBtn1.get(i1).click();
								break;}
						}}
						catch(Exception e) {Thread.sleep(1000);}

//REDIO BUTTON
					try {
						
						String value = celldata.getCell(3).getStringCellValue();
						Thread.sleep(1000);
						for(int i1=0;i1<=allRedio.size();i1++) {
				//			System.out.println("List of Secound Redio Buttons ="+allRedio.get(i1).getText());
						if(allRedio.get(i1).getText().equalsIgnoreCase(value)) {
							allRedio.get(i1).click();
							break;}
					}}
					catch(Exception e) {Thread.sleep(1000);}

//REDIO BUTTON 3
					try {
						String  option= celldata.getCell(4).getStringCellValue();
						Thread.sleep(1000);
						for(int i1=0;i1<=allRedioBtn3.size();i1++) {
				//			System.out.println("List of Secound Redio Buttons ="+allRedioBtn3.get(i1).getText());
						if(allRedioBtn3.get(i1).getText().equalsIgnoreCase(option)) {
							allRedioBtn3.get(i1).click();
							break;}
					}}
					catch(Exception e) {Thread.sleep(1000);}
					
//OFFICE
					try {
						String  office1= celldata.getCell(5).getStringCellValue();
						Select se1=new Select(office);
						se1.selectByVisibleText(office1);}
					catch(Exception e) {Thread.sleep(800);}
					
//PARTY
					try {
						Thread.sleep(1000);
						String  party= celldata.getCell(6).getStringCellValue();
						partD.click();
						partyText.sendKeys(party);
						partyText.sendKeys(Keys.ENTER);
						for(int i1=2;i1<=allParty.size();i1++)
						{
						if(allParty.get(i1).getText().equalsIgnoreCase(party)) {
							allParty.get(i1).click();
							break;}}
					}					
					catch(Exception a) {Thread.sleep(1000);}
					
//PARTY OU
					try {
						String ou= celldata.getCell(7).getStringCellValue();
							Select se2=new Select(partyOU);
							se2.selectByVisibleText(ou);}
					catch(Exception p) {Thread.sleep(1000);}
					
//OK BUTTON
					Thread.sleep(1500);
					try {
						okBtn.click();
						Thread.sleep(1000);
						driver.switchTo().alert().accept();
						
					}catch(Exception q) {Thread.sleep(1000);}
				}
				}
					catch(NullPointerException e) {
						Thread.sleep(500);}
				}
}
	
	
	public void basicDetail(int excel) throws Exception {
		
		  FileInputStream file5=new FileInputStream("C:\\Users\\Admin\\eclipse-workspace\\Freight_Journal\\Test_data\\Data1.xlsx");	
			
			
				XSSFWorkbook workbook=new XSSFWorkbook(file5);
				XSSFSheet sheet = workbook.getSheet("basicDetail");
				int rowcount = sheet.getLastRowNum();
				int colcount = sheet.getRow(7).getLastCellNum();
				System.out.println("basicDetail rowcount:"+rowcount+"basicDetail colcount"+colcount);

				for(int i=7;i<=rowcount;i++)
				{
					XSSFRow celldata = sheet.getRow(i);	
					try {
					int idNo = (int) celldata.getCell(1).getNumericCellValue();
					
					if(idNo == excel) {
						
//PLACE OF SUPPLY
						try {
							String val1 = celldata.getCell(2).getStringCellValue();
							placeOfSupply.click();
							placeOfSupplyText.sendKeys(val1);
							placeOfSupplyText.sendKeys(Keys.ENTER);
							for(int i1=2;i1<=allPlaceOfSupply.size();i1++) {
							if(allPlaceOfSupply.get(i1).getText().equalsIgnoreCase(val1)) {
								allPlaceOfSupply.get(i1).click();
								break;}
						}}
						catch(Exception e) {Thread.sleep(1000);}

//SERVICE ACCOUNT TYPE
						try {
							String val = celldata.getCell(3).getStringCellValue();
							Select se2=new Select(serviceAccType);
							se2.selectByVisibleText(val);}
						catch(Exception e) {Thread.sleep(1000);}
						
//PARTY ACCOUNT TYPE
						try {
							String val2 = celldata.getCell(4).getStringCellValue();
							Select se2=new Select(partyAccType);
							se2.selectByVisibleText(val2);}
						catch(Exception a) {Thread.sleep(1000);}
						
//PARTY OU
						try {
							String val2 = celldata.getCell(5).getStringCellValue();
							partyouD.click();
							Select se3=new Select(partuOU2);
							se3.selectByVisibleText(val2);
							Thread.sleep(1000);
							okBtn2.click();
							Thread.sleep(1000);
							driver.switchTo().alert().accept();}
						catch(Exception w) {Thread.sleep(1000);}
						
//TAX TYPE
						try {
							String val3 = celldata.getCell(6).getStringCellValue();
							Select se4=new Select(taxType);
							se4.selectByVisibleText(val3);}
						catch(Exception r) {Thread.sleep(1000);}
						
// DATE
						try {
							String tranDate = celldata.getCell(7).getStringCellValue();
							String duDate = celldata.getCell(8).getStringCellValue();
							String date = celldata.getCell(9).getStringCellValue();
							
							transDate.clear();
							JavascriptExecutor js2=(JavascriptExecutor)driver;
							js2.executeScript("arguments[0].value='"+ tranDate +"'" , transDate);
							
							Thread.sleep(1000);
							docDate.clear();		
							JavascriptExecutor js4=(JavascriptExecutor)driver;
							js4.executeScript("arguments[0].value='"+ date +"'" , docDate);
							Thread.sleep(1500);
						
							Thread.sleep(1000);
							dueDate.clear();
							JavascriptExecutor js5=(JavascriptExecutor)driver;
							js5.executeScript("arguments[0].value='"+ duDate +"'" , dueDate);
						}
							
						catch(Exception d) {Thread.sleep(1000);}
						
						
//ARAP REF
						try {
							String ref = celldata.getCell(10).getStringCellValue();
							arpRef.sendKeys(ref);}
						catch(Exception p) {Thread.sleep(1000);}
						
//INVOICE TYPE
						try {
							Thread.sleep(1000);
							String invoice = celldata.getCell(11).getStringCellValue();
							Select se=new Select(invoiceType);
							se.selectByVisibleText(invoice);
							Thread.sleep(800);}
						catch(Exception d) {Thread.sleep(1000);}
						
						
				}
				}
					catch(NullPointerException e) {
						Thread.sleep(500);}
					
				}}
	
	
public void addDetails(int excel) throws Exception {
	
	FileInputStream file5=new FileInputStream("C:\\Users\\Admin\\eclipse-workspace\\Freight_Journal\\Test_data\\Data1.xlsx");	
	
	
	XSSFWorkbook workbook=new XSSFWorkbook(file5);
	XSSFSheet sheet = workbook.getSheet("addDetails");
	int rowcount = sheet.getLastRowNum();
	int colcount = sheet.getRow(7).getLastCellNum();
	System.out.println("addDetails rowcount:"+rowcount+"addDetails colcount"+colcount);

	for(int i=7;i<=rowcount;i++)
	{
		XSSFRow celldata = sheet.getRow(i);	
		try {
		int idNo = (int) celldata.getCell(1).getNumericCellValue();
		
		if(idNo == excel) {
			
//ADD DETAILS
			try {
				Thread.sleep(1000);
				addDetailsBtn.click();}
			catch(Exception s) {Thread.sleep(1000);}
			
//BOOK
			try {
				bookD.click();
				String bookName = celldata.getCell(2).getStringCellValue();
				bookText.sendKeys(bookName);
				Thread.sleep(1000);
				bookText.sendKeys(Keys.ENTER);
				for(int i2=2;i2<=allBookRef.size();i2++) {
					if(allBookRef.get(i2).getText().equalsIgnoreCase(bookName)) {
						allBookRef.get(i2).click();}}	
			}
			catch(Exception d) {Thread.sleep(1000);}
			
//TARIFF
		
			try {
				tariffD.click();
				String tarif = celldata.getCell(3).getStringCellValue();
				Thread.sleep(1000);
				tariffText.sendKeys(tarif);
				tariffText.sendKeys(Keys.ENTER);
				for(int a=2;a<=tariffList.size();a++) {
					if(tariffList.get(a).getText().equalsIgnoreCase(tarif)) {
						tariffList.get(a).click();}}
			}
			catch(Exception e) {Thread.sleep(1000);}
			
//DESCRIPTION	
			
			try {
				String description = celldata.getCell(4).getStringCellValue();
				desc.clear();
				desc.sendKeys(description);}
			catch(Exception f) {Thread.sleep(800);}
			
//UOM
			
			try {
				String uomA = celldata.getCell(5).getStringCellValue();
				Select se5=new Select(uom);
				se5.selectByVisibleText(uomA);
			}
			catch(Exception f) {Thread.sleep(1000);}
			
//RATE
			Thread.sleep(800);
			try {	
		rate.click();
        double rateT = celldata.getCell(6).getNumericCellValue();
		JavascriptExecutor js4=(JavascriptExecutor)driver;
		js4.executeScript("arguments[0].value='"+ rateT +"'" , rate);}
		catch(Exception a) {Thread.sleep(1000);}

//QTY
			Thread.sleep(1000);
			try {
			double qty1 = celldata.getCell(7).getNumericCellValue();
			JavascriptExecutor js3=(JavascriptExecutor)driver;
			js3.executeScript("arguments[0].value='"+ qty1 +"'" , qty);
			Thread.sleep(1000);
			driver.findElement(By.xpath("//span[@id='ctl00_ContentPlaceHolder1_Label41115']")).click();}
			catch(Exception a) {Thread.sleep(1000);}
		
//CURRENCY
			Thread.sleep(1000);
			try {
				String curr = celldata.getCell(8).getStringCellValue();
				Select se1=new Select(currency);
				se1.selectByVisibleText(curr);
				Thread.sleep(1000);}
			catch(Exception a) {Thread.sleep(1000);}
			
//REMARK
			Thread.sleep(1000);
			try {
				String rem = celldata.getCell(9).getStringCellValue();
				remark.clear();
				remark.sendKeys(rem);}
			catch(Exception r) {Thread.sleep(1000);}
			
//ADD
			Thread.sleep(1500);
			try {
				addBtn.click();
//Drsft Succ Save      * Document is saved as Draft Successfully
				if(succMasg.getText().contains("* Document is saved as Draft Successfully")) {
					Thread.sleep(1000);
					closeBtn.click();
				}}
			catch(Exception g) {Thread.sleep(1000);}
		
//Final Remark
			
			JavascriptExecutor jse = (JavascriptExecutor)driver;
			jse.executeScript("window.scrollBy(0,300)");
			
			Thread.sleep(1000);
			String fRem = celldata.getCell(10).getStringCellValue();
			finalRemark.sendKeys(fRem);
			
			
	}
	}
		catch(NullPointerException e) {
			Thread.sleep(500);}
		
	}}


public void save() throws Exception {
	try {
		saveBtn.click();
		Thread.sleep(1000);
		driver.switchTo().alert().accept();
		Thread.sleep(2000);
		try {
			if(driver.switchTo().alert().getText().contains("For selected Party ,same APRef is used for another Document"))
				driver.switchTo().alert().accept();}
		catch(Exception e) {Thread.sleep(1000);}
		
	//	System.out.println("Massage ="+jurnalSave.getText());                 //* Freight Journal 'CHN/GPJ/00007/23-24'; Voucher Number 'NWL-PJ-000008-23-24' Saved Successfully
		if((jurnalSave.getText().contains("* Freight Journal ")) || (jurnalSave.getText().contains(" Voucher Number ")) || (jurnalSave.getText().contains(" Saved Successfully"))) {
			driver.findElement(By.xpath("//a[text()='Close']")).click();
			Thread.sleep(1500);
			close.click();
			
			Thread.sleep(1000);
			 undo.click();
			 
			 Thread.sleep(1000);
			 newBtnS.click();	
		}
		
		
		
	}
	catch(Exception s) {
		driver.switchTo().alert().getText();
	Thread.sleep(1500);
	driver.findElement(By.xpath("//a[text()='Close']")).click();
	   Thread.sleep(1000);
	//   close.click();
	   undo.click();
		 
	   Thread.sleep(1000);
	   newBtnS.click();
	   }
	
}


}
