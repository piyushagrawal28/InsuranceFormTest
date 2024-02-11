package carInsuranceRegistration;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.time.Duration;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class CarInsuranceTest {

	WebDriver driver;
	String url = "https://sampleapp.tricentis.com/101/app.php";
	static String testDataPath = "src/test/resources/testdata/InsuranceFormData.xlsx";
	static XSSFWorkbook wrkbk;
	static XSSFSheet sheet;
	static XSSFRow row;
	static XSSFCell column;
	static String sheetName = "InsuranceFormData";
	static FileInputStream fs;
	static DataFormatter dataFormate = new DataFormatter();
	
	@BeforeClass
	void setupBrowser()
	{
		driver = new ChromeDriver();
	}
	@DataProvider(name="FormData")
	public static Object[][] readTestData()
	{
		try {
			fs = new FileInputStream(testDataPath);
			wrkbk = new XSSFWorkbook(fs);
			int shtcount = wrkbk.getNumberOfSheets();
			Object data[][] = new Object[shtcount][38];
			for(int i=1; i<=shtcount; i++)
			{
				sheet = wrkbk.getSheet(sheetName + i);
				for(int j=1; j<=38; j++) {
					row = sheet.getRow(j);
					if(row==null)
					{
						data[i-1][j-1]="";
					}
					else {
						column = row.getCell(1);
						if(column==null)
						{
							data[i-1][j-1]="";
						}
						else {
						String value = dataFormate.formatCellValue(column);
						data[i-1][j-1]=value;
						}
					}
				}	
			}
			return data;
		}catch(Exception e) {
			System.out.print("Unable to read data file : " + e);
			return null;
		}
	}
	
	@Test(dataProvider = "FormData")
	void insuranceRegistration(String values[])
	{
		driver.get(url);
		driver.manage().window().maximize();
		driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(30));
		Select make = new Select(driver.findElement(By.id("make")));
		make.selectByValue(values[0]);
		Select model = new Select(driver.findElement(By.id("model")));
		model.selectByValue(values[1]);
		driver.findElement(By.id("cylindercapacity")).sendKeys(values[2]);
		driver.findElement(By.id("engineperformance")).sendKeys(values[3]);
		driver.findElement(By.id("dateofmanufacture")).sendKeys(values[4].substring(1));
		Select numberOfSeats = new Select(driver.findElement(By.id("numberofseats")));
		numberOfSeats.selectByValue(values[5]);
		driver.findElement(By.xpath("//span[preceding-sibling::input[@id='righthanddrive"+values[6].toLowerCase()+"']]")).click();
		Select numberofseatsmotorcycle = new Select(driver.findElement(By.id("numberofseatsmotorcycle")));
		numberofseatsmotorcycle.selectByValue(values[7]);
		Select fuelType = new Select(driver.findElement(By.id("fuel")));
		fuelType.selectByValue(values[8]);
		driver.findElement(By.id("payload")).sendKeys(values[9]);
		driver.findElement(By.id("totalweight")).sendKeys(values[10]);
		driver.findElement(By.id("listprice")).sendKeys(values[11]);
		driver.findElement(By.id("licenseplatenumber")).sendKeys(values[12]);
		driver.findElement(By.id("annualmileage")).sendKeys(values[13]);
		driver.findElement(By.id("nextenterinsurantdata")).click();
		// Insurance User Details Page
		driver.findElement(By.id("firstname")).sendKeys(values[14]);
		driver.findElement(By.id("lastname")).sendKeys(values[15]);
		driver.findElement(By.id("birthdate")).sendKeys(values[16].substring(1));
		driver.findElement(By.xpath("//span[preceding-sibling::input[@id='gender" + values[17].toLowerCase()+"']]")).click();
		driver.findElement(By.id("streetaddress")).sendKeys(values[18]);
		Select country = new Select(driver.findElement(By.id("country")));
		country.selectByValue(values[19]);
		driver.findElement(By.id("zipcode")).sendKeys(values[20]);
		driver.findElement(By.id("city")).sendKeys(values[21]);
		Select occupation = new Select(driver.findElement(By.id("occupation")));
		occupation.selectByValue(values[22]);
		String arg[] = values[23].split(",");
		for(int i=0; i<arg.length; i++)
		{
			driver.findElement(By.xpath("//span[preceding-sibling::input[@id='" + arg[i].trim().toLowerCase().replace(" ", "") +"']]")).click();
		}
		driver.findElement(By.id("website")).sendKeys(values[24]);
		driver.findElement(By.id("open")).click();
		try {
		Robot rbt = new Robot();
		Thread.sleep(3000);
		driver.switchTo().activeElement().sendKeys(values[25]);
		rbt.keyPress(KeyEvent.VK_ALT);
		rbt.keyPress(KeyEvent.VK_F4);
		rbt.keyRelease(KeyEvent.VK_F4);
		rbt.keyRelease(KeyEvent.VK_ALT);
		Thread.sleep(3000);
		}catch(Exception e)
		{
			System.out.println("Robot class is not working in the test enviroment : " + e);
		}
		driver.findElement(By.id("nextenterproductdata")).click();
		// Product Details Page
		driver.findElement(By.id("startdate")).sendKeys(values[26].substring(1));
		Select insuranceSum = new Select(driver.findElement(By.id("insurancesum")));
		insuranceSum.selectByValue(values[27]);
		Select meritRating = new Select(driver.findElement(By.id("meritrating")));
		meritRating.selectByValue(values[28]);
		Select damageInsurance = new Select(driver.findElement(By.id("damageinsurance")));
		damageInsurance.selectByValue(values[29]);
		String argu[] = values[30].split(",");
		for(int i=0; i<argu.length; i++)
		{
			driver.findElement(By.xpath("//span[preceding-sibling::input[@id='" + argu[i].trim().replace(" ", "")+"']]")).click();
		}
		Select courtesyCar = new Select(driver.findElement(By.id("courtesycar")));
		courtesyCar.selectByValue(values[31]);
		driver.findElement(By.id("nextselectpriceoption")).click();
		// Price Details Page 
		driver.findElement(By.xpath("//span[preceding-sibling::input[@id='select" + values[32].toLowerCase() + "']]")).click();
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		WebElement nextSendQuotebutton = driver.findElement(By.id("nextsendquote"));
		wait.until(ExpectedConditions.elementToBeClickable(nextSendQuotebutton));
		nextSendQuotebutton.click();
		// Quotation Mailing Details
		driver.findElement(By.id("email")).sendKeys(values[33]);
		driver.findElement(By.id("phone")).sendKeys(values[34]);
		driver.findElement(By.id("username")).sendKeys(values[35]);
		driver.findElement(By.id("password")).sendKeys(values[36]);
		driver.findElement(By.id("confirmpassword")).sendKeys(values[36]);
		driver.findElement(By.id("Comments")).sendKeys(values[37]);
		driver.findElement(By.id("sendemail")).click();
		// Verifying success alert
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.sa-icon.sa-success")));
		WebElement heading = driver.findElement(By.xpath("//h2[preceding-sibling::div[preceding-sibling::div[contains(@class,'sa-icon sa-success')]]]"));
		String successMsg = heading.getText();
		String expectedMsg = "Sending e-mail success";
		Assert.assertTrue(successMsg.contains(expectedMsg),"Insurance form successfully submitted");
		driver.findElement(By.className("confirm")).click();
				
	}
	@AfterClass
	void tearDown()
	{
		driver.close();
	}
}
