package org.base;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {

	public static WebDriver driver;
	public static Robot ro;
	public static JavascriptExecutor js;
	public static Actions ac;

	public static void lunchbrowser() {

		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().window().maximize();

	}

	public static void lunchurl(String text) {
		driver.get(text);
	}

	public static void filltext(WebElement e, String text) {
		e.sendKeys(text);

	}

	public static void clickbtn(WebElement e) {

		e.click();

	}

	public static void dropdown_visabletext(WebElement e, String text) {
		Select sc = new Select(e);

		sc.selectByVisibleText(text);

	}

	public static void dropdown_index(WebElement e, int indexno) {
		Select sc = new Select(e);

		sc.selectByIndex(indexno);

	}

	public static void dropdown_value(WebElement e, String text) {
		Select sc = new Select(e);

		sc.selectByValue(text);

	}

	public static void windowhandels(int No) {
		Set<String> allw = driver.getWindowHandles();

		List<String> pk = new LinkedList<String>();

		pk.addAll(allw);

		driver.switchTo().window(pk.get(No));

	}

	public static void windowhandles_method2() {
		String parent = driver.getWindowHandle();

		Set<String> child = driver.getWindowHandles();
		for (String x : child) {
			if (!parent.equals(x)) {
				driver.switchTo().window(x);

			}
		}

	}

	public static void timenote() {
		Date a = new Date();
		System.out.println(a);

	}

	public static void screenShot(String Filepath, String filename) throws IOException {
		TakesScreenshot ts = (TakesScreenshot) driver;
		File screenshotAs = ts.getScreenshotAs(OutputType.FILE);
		File fil = new File(Filepath + filename + ".png");
		FileUtils.copyFile(screenshotAs, fil);

	}

	public static void rightclk(WebElement text) {
		ac = new Actions(driver);

		ac.contextClick(text).build().perform();
	}

	public static void dubleclk(WebElement e) {
		ac = new Actions(driver);
		ac.doubleClick(e).build().perform();

	}

	public static void mouseover(WebElement e) {
		ac = new Actions(driver);
		ac.moveToElement(e).build().perform();

	}

	public static void keydown(WebElement e) throws AWTException {
		ro = new Robot();

		ro.keyPress(KeyEvent.VK_DOWN);
		ro.keyRelease(KeyEvent.VK_DOWN);

	}

	public static void keyup(WebElement e) throws AWTException {
		ro = new Robot();

		ro.keyPress(KeyEvent.VK_UP);
		ro.keyRelease(KeyEvent.VK_UP);
	}

	public static void copy(WebElement e) throws AWTException {
		ro = new Robot();

		ro.keyPress(KeyEvent.VK_CONTROL);
		ro.keyPress(KeyEvent.VK_C);

		ro.keyRelease(KeyEvent.VK_CONTROL);
		ro.keyRelease(KeyEvent.VK_C);

	}

	public static void paste(WebElement e) throws AWTException {
		ro = new Robot();

		ro.keyPress(KeyEvent.VK_CONTROL);
		ro.keyPress(KeyEvent.VK_V);

		ro.keyRelease(KeyEvent.VK_CONTROL);
		ro.keyRelease(KeyEvent.VK_V);

	}

	public static void enterkey() throws AWTException {
		ro = new Robot();
		ro.keyPress(KeyEvent.VK_ENTER);
		ro.keyRelease(KeyEvent.VK_ENTER);

	}

	public static void keytab() throws AWTException {
		Robot ro = new Robot();
		ro.keyPress(KeyEvent.VK_TAB);
		ro.keyRelease(KeyEvent.VK_TAB);

	}

	public static void closewindow() {
		driver.close();

	}

	public static void quitwindow() {
		driver.quit();

	}

	public static void java_click(WebElement a) {
		js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].click()", a);

	}

	public static void Java_filltext(String text, WebElement a) {

		js = (JavascriptExecutor) driver;

		js.executeScript("arguments[0].setAttribute('value','" + text + "')", a);
	}

	public static void navigat_refresh() {
		driver.navigate().refresh();

	}

	public static void navigat_forword() {
		driver.navigate().forward();

	}

	public static void navigat_backword() {
		driver.navigate().back();

	}

	public static void implecitywait() {
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.MINUTES);

	}

	public static void explicitywait(int seconds, int milliesec) {
		FluentWait<WebDriver> f = new FluentWait<WebDriver>(driver).withTimeout(Duration.ofSeconds(seconds))
				.pollingEvery(Duration.ofMillis(milliesec)).ignoring(Throwable.class);

	}

	public static void createxcelsheet(int row, int cell) throws IOException {
		File f = new File("C:\\Users\\WELCOME\\eclipse-workspace\\Baseclass\\src\\test\\java\\data\\Facebook.xlsx");

		Workbook w = new XSSFWorkbook();

		Sheet sheet = w.createSheet("insta");
		Row createRow1 = sheet.createRow(row);
		Cell createCell1 = createRow1.createCell(cell);

	}

	public static String excelread(String Filetext, String sheetText, int row, int cell) throws IOException {

		FileInputStream f = new FileInputStream(
				"C:\\Users\\WELCOME\\eclipse-workspace\\Baseclass\\src\\test\\java\\data\\" + Filetext + ".xlsx");

		Workbook workbook = new XSSFWorkbook(f);

		Sheet sheet = workbook.getSheet(sheetText);
		Row row2 = sheet.getRow(row);
		Cell cell2 = row2.getCell(cell);

		int cellType = cell2.getCellType();

		String value;
		if (cellType == 1) {
			value = cell2.getStringCellValue();

		} else if (DateUtil.isCellDateFormatted(cell2)) {

			Date d = cell2.getDateCellValue();
			SimpleDateFormat s = new SimpleDateFormat("dd-MM-yyyy");
			value = s.format(d);

		} else {
			double a = cell2.getNumericCellValue();
			long l = (long) a;

			value = String.valueOf(l);

		}
		return value;

	}

	public static void excelwrite(String Filename, String sheettext, int row, int cel, String setvalue)
			throws IOException {

		File f = new File(
				"C:\\Users\\WELCOME\\eclipse-workspace\\Baseclass\\src\\test\\java\\data" + Filename + ".xlsx");
		Workbook w = new XSSFWorkbook();

		Sheet sheet = w.getSheet(sheettext);
		Row row2 = sheet.getRow(row);
		Cell cell2 = row2.getCell(cel);

		cell2.setCellValue(setvalue);

		FileOutputStream f2 = new FileOutputStream(f);
		w.write(f2);

	}

	public static void excelTestData_printallvalue(String sheettext, String Filename) throws IOException {

		FileInputStream fl = new FileInputStream(
				"C:\\Users\\WELCOME\\eclipse-workspace\\Baseclass\\src\\test\\java\\data\\" + Filename + ".xlsx");

		Workbook wo = new XSSFWorkbook(fl);

		Sheet sheet = wo.getSheet(sheettext);

		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);

			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);

				int alltype = cell.getCellType();
				String value;
				if (alltype == 1) {
					value = cell.getStringCellValue();

				} else if (DateUtil.isCellDateFormatted(cell)) {

					Date dd = cell.getDateCellValue();
					SimpleDateFormat s = new SimpleDateFormat("dd-mm-yyyy");
					value = s.format(dd);

				} else {
					double nu = cell.getNumericCellValue();

					long a = (long) nu;

					value = String.valueOf(a);

				}
				System.out.println(value);

			}

		}

	}

}
