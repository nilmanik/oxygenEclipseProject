package test;



import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Test1 {
	
	private static XSSFSheet ExcelWSheet;
	 
	private static XSSFWorkbook ExcelWBook;

	private static XSSFCell Cell;

	private static XSSFRow Row;
	
	public WebDriver driver;
	public WebDriverWait wait;
	String appURL ="https://mail.google.com";
	
	@BeforeMethod
	public void setup(){
		System.setProperty("webdriver.chrome.driver", "C:\\softwares\\selenium3\\chromedriver_win32\\chromedriver.exe");
		driver=new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.get(appURL);
	}
	
	@Test(dataProvider="empLogin")
	public void test(String username,String password) throws InterruptedException{
		
		//driver.findElement(By.xpath("//a[@class='login']")).click();
		driver.findElement(By.id("identifierId")).sendKeys(username);
		driver.findElement(By.xpath("//span[contains(text(),'Next')]")).click();
		//Thread.sleep(6000);
		driver.findElement(By.name("password")).sendKeys(password);
		Thread.sleep(7000);
		driver.findElement(By.xpath("//span[contains(text(),'Next')]")).click();
		/*Thread.sleep(7000);
		//driver.findElement(By.xpath("//span[text()='Yes']")).click();
		//Thread.sleep(7000);
		driver.findElement(By.xpath("//span[@class='gb_db gbii']")).click();
		//Thread.sleep(10000);
		driver.findElement(By.xpath("//a[@id='gb_71']")).click();*/

}
	
	
	
	@DataProvider(name="empLogin")
	public Object[][] loginData() {
		Object[][] arrayObject = getExcelData("C:\\sampledoc.xlsx","Sheet1");
		return arrayObject;
	}
	/**
	 * @param File Name
	 * @param Sheet Name
	 * @return
	 */
	public String[][] getExcelData(String fileName, String sheetName) {
		String[][] arrayExcelData = null;
		try {
			FileInputStream fs = new FileInputStream(fileName);
			ExcelWBook = new XSSFWorkbook(fs);
			ExcelWSheet = ExcelWBook.getSheet(sheetName);
			int totalNoOfRows = ExcelWSheet.getLastRowNum();
			int totalNoOfCols = ExcelWSheet.getRow(0).getLastCellNum();
			
			
			arrayExcelData = new String[totalNoOfRows-1][totalNoOfCols];
			
			for (int i= 1 ; i < totalNoOfRows; i++) {

				for (int j=0; j < totalNoOfCols; j++) {
					arrayExcelData[i-1][j] = getCellData(i, j);
				}

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return arrayExcelData;
	}
	public static String getCellData(int RowNum, int ColNum) throws Exception {
		 
		try{

			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);

			int dataType = Cell.getCellType();

			if  (dataType == 3) {

				return "";

			}else{

				String CellData = Cell.getStringCellValue();

				return CellData;

			}}
			catch (Exception e){

			System.out.println(e.getMessage());

			throw (e);

			}}
	@AfterMethod
	public void endTest(){
		driver.close();
	}
}