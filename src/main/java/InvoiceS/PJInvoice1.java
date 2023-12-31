package InvoiceS;

import java.util.NoSuchElementException;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import Base.Browser;
import Utility.CommonFile;

public class PJInvoice1 extends Browser{
	
	@FindBy(xpath = "//input[@title='Username']")
	private WebElement username;
	@FindBy(xpath = "//input[@title='Password']")
	private WebElement pass;
	@FindBy(xpath = "//button[text()='Login']")
	private WebElement loginBtn;

	@FindBy(xpath = "//i[@class='nav-icon fa fa-forward']")
	private WebElement IFFBtn;
	@FindBy(xpath = "//span[text()='IFF Financials']")
	private WebElement financeBtn;
	@FindBy(xpath = "//span[text()='Freight Journal (PJ)']")
	private WebElement pjInvoice;

	public PJInvoice1() {
		PageFactory.initElements(driver, this);
	}

	public void verifyLoginApp() throws Exception {
		username.sendKeys(readExcelFileFinal(3, 3));
		pass.sendKeys(readExcelFileFinal(3, 4));
		Thread.sleep(1000);
		loginBtn.click();

	}

	public void verifyIFFBtn() throws Exception {
		try {
			Thread.sleep(1000);
		IFFBtn.click();}
		catch(NoSuchElementException e) {
		}
	}

	public void verifyFinanceBtn() throws Exception {
		Thread.sleep(800);
		financeBtn.click();
	}

	public void verifyFreightJournalBtn() throws Exception {
		pjInvoice.click();
	}
	
	

}
