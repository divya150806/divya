package baseclass;

import java.awt.AWTException;
import java.awt.Robot;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class Baseclass {
	public static WebDriver driver;
	Actions action;
	Alert alert;
	JavascriptExecutor executor;
	Select select;
	//1
	public static void browserLaunch(String url)
	{
		driver=new ChromeDriver();
		driver.get(url);
		driver.manage().window().maximize();	 
	}
	
	public static void browserLaunchMS(String URL)
	{
		driver=new EdgeDriver();
		driver.get("URL");
		driver.manage().window().maximize();
		
	}
	//2 StringID
	public static WebElement findelementId(String Id)
	{
		WebElement elementId=driver.findElement(By.id(Id));
		return elementId;	
	}

	//3 StringName
	public static WebElement findelementName(String Name)
	{
		WebElement elementName=driver.findElement(By.name(Name));

		return elementName;
	}
	//4 sendkeys

	public static void sendkeyvalues(WebElement element, String data)
	{
		element.sendKeys(data);
	}

	//5xpath

	public static WebElement findXpath(String xpath)
	{
		WebElement elementXpath=driver.findElement(By.xpath(xpath));
		return elementXpath;

	}
	//6 close browser
	public static void closeBrowser()
	{
		driver.close();
	}
	// 7 Action class
	public void actionClass(WebElement element)
	{
		action=new Actions(driver);
		action.moveToElement(element).perform();

	}

	//8 Drag and Drop
	public void dragAndDrop(WebElement source, WebElement Dest)
	{

		action.dragAndDrop(source, Dest).perform();


	}

	//9 double click
	public void doubleClick(WebElement Data )
	{

		action.doubleClick(Data).perform();

	}

	//10 Context click--Right click

	public void contextClick(WebElement args)
	{

		action.contextClick(args).perform();
	}

	//11 Robot Class
	public void robotClass(int keypress,int keyrelease) throws AWTException
	{
		Robot robot=new Robot();
		robot.keyPress(keypress);
		robot.keyRelease(keyrelease);

	}

	// 12 to accept the alerts

	public void alertsAcccept()
	{
		alert=driver.switchTo().alert();
		alert.accept();

	}
	
	// 13.to Reject the alert
	public void alertsReject()
	{
		alert=driver.switchTo().alert();
		alert.dismiss();
	}

	//14 Get the String of the element

	public WebElement getText(WebElement textelement)
	{
		textelement.getText();
		return textelement;
	}
	//15 JavaScrip Executor

	public WebElement javaScriptExecutor(int args,String data,WebElement element)
	{
		JavascriptExecutor executor=(JavascriptExecutor) driver;
		Object	text=executor.executeScript("arguments["+ args +"].setAttribute('value',"+ data+")", element);
		return element;

	}
	//16 scrollDown
	public WebElement scrollDown(int num,WebElement element)
	{
		executor.executeScript("arguments["+num+"].scrollIntoView(true)", element);
		return element;
	}
	//17 scrollUp
	public WebElement scrollUp(int num,WebElement element)
	{
		executor.executeScript("arguments["+num+"].scrollIntoView(false)", element);
		return element;
	}

	//18 Screenshot

	public static void screenShot(String filepath) throws IOException
	{
		TakesScreenshot screenshot = (TakesScreenshot) driver;

		File sourceFile = screenshot.getScreenshotAs(OutputType.FILE);

		File dest = new File(filepath);

		FileUtils.copyFile(sourceFile, dest);
	}

	//19 DropDownClass-Selectbyindex

	public void dropdownIndex(WebElement element,int index )
	{
		Select select=new Select(element);
		select.selectByIndex(index);

	}
	//20 dropdown -SelectBy Value
	public void dropdownValue(WebElement element,String name)
	{
		Select select=new Select(element);
		select.selectByValue(name);

	}
	//21 Select By VisibleText

	public void dropdownText(WebElement element,String text )
	{
		Select select=new Select(element);
		select.selectByVisibleText(text);

	}
	//22 how to get list of elements from the dropdown

	public void listOfelemnts(WebElement element) {


		select = new Select(element);

		List<WebElement> options = select.getOptions();
		System.out.println(options.size());
		for (int i = 0; i < options.size(); i++) {
			WebElement webElement = options.get(i);
			System.out.println(webElement.getText());
		}
	}
	//23 how to do the multi select

	public void multiSelect(WebElement element)
	{
		select=new Select(element) ;
		List<WebElement> allselectedoptions = select.getAllSelectedOptions();
		for (WebElement option : allselectedoptions)
		{
			System.out.println("Selected Option: " + option.getText());
			option.click();
		}
	}
	//24 Frames

	public void frames(WebElement element1, WebElement element2 )
	{
		driver.switchTo().frame(element1);

		driver.switchTo().defaultContent();
		driver.switchTo().frame(element2);

	}

	//25 frames using index
	public void frames(int index1, int index2 )
	{
		driver.switchTo().frame(index1);

		driver.switchTo().defaultContent();
		driver.switchTo().frame(index2);

	}

	//26 Windows handling

	public void windowsHandling(WebElement element)
	{

		// to get the parentWindow ID
		String parentWindow= driver.getWindowHandle();
		System.out.println(element);
		//to get all WindlowsID
		Set<String> allwindow=driver.getWindowHandles();
		System.out.println(allwindow);
		for (String eachwindow : allwindow)
		{
			if (!parentWindow.equals(eachwindow))
			{
				driver.switchTo().window(eachwindow);
			}
		}
	}


	//27 Excel read

	public String getDatafromExcel(String location,String sheetName,int rowNo,int cellNo) throws IOException
	{
		String cellData=null;
		File file = new File(location);
		FileInputStream inputstream=new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(inputstream);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(rowNo);
		Cell cell = row.getCell(cellNo);
		CellType cellType = cell.getCellType();

		if(cellType==CellType.STRING)
		{
			cellData = cell.getStringCellValue();
		}
		if (cellType == CellType.NUMERIC)
		{
			if (DateUtil.isCellDateFormatted(cell))
			{
				Date dateCellValue = cell.getDateCellValue();
				SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
				cellData = dateFormat.format(dateCellValue);
			}
			else
			{
				double numericCellValue = cell.getNumericCellValue();
				long l = (long)numericCellValue;
				cellData=String.valueOf(l);

			}

		}
		return cellData;		

	}

	//	28 Locators--class
	public WebElement findclassName(String classname)
	{
		WebElement elementName=driver.findElement(By.className(classname));

		return elementName;
	}
	
	// 29 Method to get the pagetitle
	public String pageTitle() {
		String title = driver.getTitle();
		return title;
		
	}
	//30 Method to get the currentURL
	
	public String currentURL()
	{
		 return driver.getCurrentUrl();
	}
	
	//31 Navigate to a URL
	
	public void navigateURL(String url)
	{
		 driver.navigate().to(url);
	}
	
	
	//32 Navigate to forward
	
	public void navigateforward()
	{
		driver.navigate().forward();
	}
	
	//34 Navigate to backward
	public void navigateBack()
	{
		driver.navigate().back();
	}

	//35 to get the pagesource
	public String getPagesource()
	{
		return driver.getPageSource();
	}
	// 36 to maximize the browser
	public void maximizetheBrowser()
	{
		driver.manage().window().maximize();
	}
	//37 to minimize the browser
	public void minimizetheBrowser()
	{
		driver.manage().window().minimize();
	}
	//38 Switch to Alert and getText
	public String switchToAlertAndGetText() {
		 Alert alert = driver.switchTo().alert();
		 String alertText = alert.getText();
		 alert.accept();
		 return alertText;
		
	}
	//39 MouseHover 
	
	public void mouseHover (WebElement element)
	{
		action.moveToElement(element).perform();
	}
	
	//40  DropDownClass-deSelectbyindex

	public void deSelectbyindex(WebElement element,int index )
	{
		 select=new Select(element);
		select.deselectByIndex(index);

	}
	//41 dropdown -deSelectByValue
	public void deSelectByValue(WebElement element,String name)
	{
		 select = new Select(element);

		select.deselectByValue(name);

	}
	//42 Select By deVisibleText

	public void deVisibleText(WebElement element,String text )
	{
		 select = new Select(element);
		select.deselectByVisibleText(text);

	}
	//43 deselect All
	public void deSelectAll(WebElement element,String data)
	{
		select = new Select(element);
		select.deselectAll();
	}
	
	//44 driver quit
	
	public void quitDriver()
	{
		driver.quit();
	}
	//45 Method to perform click
	public void clickMethod(WebElement element)
	{
		element.click();
	}
	
	//46 Method to clear
	public void clearMethod(WebElement element)
	{
		element.clear();
	}
	//47 To know whether the element is displayed
	public boolean isDisplayed(WebElement element)
	{
		return element.isDisplayed();
	}
	
	//48 To know whether the element is enabled
	public boolean isEnabled(WebElement element)
	{
		return element.isEnabled();
	}
	
	//49 to get the attributes
	public String getAttributes(WebElement element,String attributeName)
	{
		return element.getAttribute(attributeName);
	}
	
	//50 Getattribute using javascrit executor
	public Object jsExecutorValues(int args,WebElement element,String attributeName)
	{
		executor = (JavascriptExecutor) driver;
		Object attributeValue = executor.executeScript(" return arguments["+ args+"].getAttribute("+ attributeName+"),"+element+" ");
		return attributeValue;
		
	}
	
	//51 Method to Enter key
	public void keyEnter(WebElement element)
	{
		element.sendKeys(Keys.ENTER);
	}
	
	//52 xpath with index
	
	/*
	 * public void xpathByIndex(String element,int index) { driver.findElement
	 * (By.xpath((element))[index]); }
	 */
	
	public Object[][] getDatafromExcel(String filePath, String sheetName) throws IOException {
		String cellData=null;
		File file = new File(filePath);
		FileInputStream inputstream=new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(inputstream);
		Sheet sheet = workbook.getSheet(sheetName);	
		 int rowCount = sheet.getPhysicalNumberOfRows();
	     int colCount = sheet.getRow(0).getPhysicalNumberOfCells();
	     
	  // Initialize a two-dimensional array to store the data
	        Object[][] data = new Object[rowCount - 1][colCount];
	        DataFormatter formatter = new DataFormatter();

	        // Loop through each row and each cell in the row
	        for (int i = 1; i < rowCount; i++) {
	            Row row = sheet.getRow(i);
	            for (int j = 0; j < colCount; j++) {
	                Cell cell = row.getCell(j);
	                if (cell != null) {
	                    // Determine the cell type and format the cell value accordingly
	                    String cellData1;
	                    if (cell.getCellType() == CellType.STRING) {
	                        cellData1 = cell.getStringCellValue();
	                    } else if (cell.getCellType() == CellType.NUMERIC) {
	                        if (DateUtil.isCellDateFormatted(cell)) {
	                            Date dateCellValue = cell.getDateCellValue();
	                            SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
	                            cellData1 = dateFormat.format(dateCellValue);
	                        } else {
	                            cellData1 = String.valueOf(cell.getNumericCellValue());
	                        }
	                    } else {
	                        cellData1 = formatter.formatCellValue(cell);
	                    }

	                    // Store the formatted cell value in the data array
	                    data[i - 1][j] = cellData1;
	                }
	            }
	        }

	        // Close the workbook and input stream
	        workbook.close();
	        inputstream.close();

	        // Return the populated data array
	        return data;
	    }
	

}

