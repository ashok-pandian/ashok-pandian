/*
	Project: Geofacets
	Copyright: Ashok Pandian / Sysvine @ 2016
	Code is to test Publisher counts for new content during every Content Release

	Last updated: 04 August, 2016
	Last content release: 04 August, 2016
 */

package com.geofacets.geoContentRT;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.StringWriter;

import static org.junit.Assert.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class PublisherStats extends Repository {
	private StringBuffer	verificationErrors	= new StringBuffer();
	String[]				pubStartYear		= new String[6], pubEndYear = new String[6];
	String[]				pubNameList			= { "Elsevier", "The Geological Society of London", "SEPM Society for Sedimentary Geology", "Society of Economic Geologists", "The Geological Society of America", "Wiley", "CNH / SENER", "American Geophysical Union", "Society of Exploration Geophysicists" };
	String[]				pubList				= { "ELS", "GSL", "SEPM", "SEG", "GSA", "Wiley", "CNH", "AGU", "SEGP" };
	String					Old, totalOldest, New, totalNewest, pubOldest, pubNewest;
	int						ASStartCell			= 0, ASEndCell = 0, PYFStartCell = 0, PYFEndCell = 0, OldestCell = 0, NewestCell = 0;
	int						mapsCountCell		= 0, docsCountCell = 0, geoCountCell = 0;
	int						Totals_row		= 4, ROW_FOR_PUBLISHER, TOTAL_PUBLISHERS = 9;
	public static String	document			= "C:\\Geofacets\\PublisherStats_DATE.xlsx";

	@BeforeTest
	public void PS_Start() throws Exception {
		//		imageBlockedFFProfile();
		//		ffProfile();
		chromeProfile();
		login("QA", "ALL");
	}

	@Test(enabled = true, groups = "PublisherStats")
	public void PS_AdvancedSearch() throws Exception {
		try {
			System.out.println("Publisher Stats starts...\n");
			System.out.println("TC1: Getting year values in Advanced Search tab...");

			driver.findElement(By.xpath("//a[@href='#advancedform']/span")).click();

			InputStream inp = new FileInputStream(document);
			Workbook wb = WorkbookFactory.create(inp);
			Sheet sheet = wb.getSheetAt(0); // 0 for first sheet
			Row row = sheet.getRow(2);

			for (int i = 0; i < row.getLastCellNum(); i++)
				if (row.getCell(i).toString().equals("Advanced Search")) {
					ASStartCell = i;
					ASEndCell = i + 1;
				}

			//Capturing and Printing Advanced Search - Total Counts
			Wait().until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//select[@id='startyear']")));

			String TotalASStart = driver.findElement(By.xpath("//select[@id='startyear']/option[@selected='selected']")).getAttribute("value");
			String TotalASEnd = driver.findElement(By.xpath("//select[@id='endyear']/option[@selected='selected']")).getAttribute("value");
			sheet.getRow(4).getCell(ASStartCell).setCellType(Cell.CELL_TYPE_NUMERIC);
			sheet.getRow(4).getCell(ASStartCell).setCellValue(Integer.parseInt(TotalASStart));
			sheet.getRow(4).getCell(ASEndCell).setCellType(Cell.CELL_TYPE_NUMERIC);
			sheet.getRow(4).getCell(ASEndCell).setCellValue(Integer.parseInt(TotalASEnd));

			//Uncheck all publisher check boxes
			for (int i = 1; i <= 9; i++)
				driver.findElement(By.xpath("//div[@id='publisherBox']/label[" + i + "]")).click();

			//Capturing and inserting date values for each publisher
			for (int i = 1; i <= 9; i++) {

				int ROW_FOR_PUBLISHER = 5 + i;
				Wait().until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//select[@id='startyear']")));

				// Checking each publisher
				driver.findElement(By.xpath("//div[@id='publisherBox']/label[contains(.,'" + pubNameList[i - 1] + "')]")).click();

				String PubASStart = driver.findElement(By.xpath("//select[@id='startyear']/option[@selected='selected']")).getAttribute("value");
				String PubASEnd = driver.findElement(By.xpath("//select[@id='endyear']/option[@selected='selected']")).getAttribute("value");
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(ASStartCell).setCellType(Cell.CELL_TYPE_NUMERIC);
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(ASStartCell).setCellValue(Integer.parseInt(PubASStart));
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(ASEndCell).setCellType(Cell.CELL_TYPE_NUMERIC);
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(ASEndCell).setCellValue(Integer.parseInt(PubASEnd));

				// Unchecking publisher
				driver.findElement(By.xpath("//div[@id='publisherBox']/label[contains(.,'" + pubNameList[i - 1] + "')]")).click();
			}

			FileOutputStream fileOut = new FileOutputStream(document);
			wb.write(fileOut);
			fileOut.close();
			System.out.println("Completed\n");

		} catch (Exception e) {

			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));
			String ex = sw.toString();
			ex = ex.substring(0, ex.indexOf("\n"));
			System.out.println(ex);
		}
	}

	@Test(enabled = true, groups = "PublisherStats")
	public void PS_OtherMethods() throws Exception {
		try {
			driver.get(baseUrl);
			System.out.println("TC2: Getting year values from content sorting, PY facet section and total counts from publishers...");

			InputStream inp = new FileInputStream(document);
			Workbook wb = WorkbookFactory.create(inp);
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(2);

			for (int i = 0; i < row.getLastCellNum(); i++)
				if (row.getCell(i).toString().equals("Publication Year Facet")) {
					PYFStartCell = i;
					PYFEndCell = i + 1;
					OldestCell = i + 2;
					NewestCell = i + 3;
					mapsCountCell = i + 4;
					geoCountCell = i + 5;
					docsCountCell = i + 6;
				}

			Wait().until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//a[@href='#searchform']/span")));
			driver.findElement(By.xpath("//a[@href='#searchform']/span")).click();

			keywSearch("\"*\"");
			Thread.sleep(2000);
			new Select(driver.findElement(By.id("sortOptions"))).selectByVisibleText("Date oldest first");
			Thread.sleep(2000);

/*
 * Get oldest & newest year values for all content through sorting
 * ------------------------------------------------------------------
 */
			String[] tempOld = driver.findElement(By.xpath("//ul[@class='display']/li[1]//font[@class='citation_d_journal']")).getText().split(",");
			Old = driver.findElement(By.xpath("//ul[@class='display']/li[1]//font[@class='citation_d_journal']")).getText();

			if (tempOld[tempOld.length - 1].contains("pp."))
				totalOldest = (Old.substring(Old.lastIndexOf(",") - 4).substring(0, 4)).trim();
			else
				totalOldest = Old.substring(Old.length() - 4, Old.length());

			new Select(driver.findElement(By.id("sortOptions"))).selectByVisibleText("Date newest first");
			Thread.sleep(2000);
			String[] tempNew = driver.findElement(By.xpath("//ul[@class='display']/li[1]//font[@class='citation_d_journal']")).getText().split(",");
			New = driver.findElement(By.xpath("//ul[@class='display']/li[1]//font[@class='citation_d_journal']")).getText();

			if (tempNew[tempNew.length - 1].contains("pp."))
				totalNewest = (New.substring(New.lastIndexOf(",") - 4).substring(0, 4)).trim();
			else
				totalNewest = New.substring(New.length() - 4, New.length());

			sheet.getRow(Totals_row).getCell(OldestCell).setCellType(Cell.CELL_TYPE_NUMERIC);
			sheet.getRow(Totals_row).getCell(OldestCell).setCellValue(Integer.parseInt(totalOldest));
			sheet.getRow(Totals_row).getCell(NewestCell).setCellType(Cell.CELL_TYPE_NUMERIC);
			sheet.getRow(Totals_row).getCell(NewestCell).setCellValue(Integer.parseInt(totalNewest));

			new Select(driver.findElement(By.id("sortOptions"))).selectByVisibleText("Date oldest first");
			Thread.sleep(2000);

/*
 * Get oldest & newest year values for all content through Pub Year facets
 * ---------------------------------------------------------------------------
 */
			driver.findElement(By.xpath("//div[@id='box5']/div/img[@id='capyr']")).click();
			Thread.sleep(1500);

			if (driver.findElement(By.xpath("//div[@id='box5']/div[@class='vmvalinks']/a[@class='view_all']")).isDisplayed())
				driver.findElement(By.xpath("//div[@id='box5']/div[@class='vmvalinks']/a[@class='view_all']")).click();
			Thread.sleep(2000);

			//	Finding no of decades (year ranges)
			int totalDecades = driver.findElement(By.xpath("//div[@id='box5']")).findElements(By.xpath("//img[starts-with(@id,'pyrParent-')]")).size();

			//	Finding no of individual years inside the range @ last
			driver.findElement(By.xpath("//img[@id='pyrParent-" + (totalDecades - 1) + "']")).click();
			Thread.sleep(2000);
			int totalSingleYears = driver.findElement(By.xpath("//div[@id='pyrParentBox" + (totalDecades - 1) + "']")).findElements(By.tagName("li")).size();

			//	Capturing and Printing 'Publisher Year Facet' - Total Counts
			String TotalPYFStart = driver.findElement(By.xpath("//input[@id='pyrnav" + (totalDecades - 1) + "_" + (totalSingleYears - 1) + "']")).getAttribute("value");
			String TotalPYFEnd = driver.findElement(By.xpath("//input[@id='pyrnav0_0']")).getAttribute("value");
			sheet.getRow(Totals_row).getCell(PYFStartCell).setCellType(Cell.CELL_TYPE_NUMERIC);
			sheet.getRow(Totals_row).getCell(PYFStartCell).setCellValue(Integer.parseInt(TotalPYFStart));
			sheet.getRow(Totals_row).getCell(PYFEndCell).setCellType(Cell.CELL_TYPE_NUMERIC);
			sheet.getRow(Totals_row).getCell(PYFEndCell).setCellValue(Integer.parseInt(TotalPYFEnd));

			// Get oldest & newest year values for all content from tab values
			Wait().until(ExpectedConditions.visibilityOfElementLocated(By.id("mapsTabLabel")));
			String TMapCount = driver.findElement(By.id("mapsTabLabel")).getText();
			String TMapCountSS = TMapCount.substring(TMapCount.indexOf("(") + 1, TMapCount.indexOf(")"));
			String TDocCount = driver.findElement(By.id("documentsTabLabel")).getText();
			String TDocCountSS = TDocCount.substring(TDocCount.indexOf("(") + 1, TDocCount.indexOf(")"));
			sheet.getRow(Totals_row).getCell(mapsCountCell).setCellType(Cell.CELL_TYPE_NUMERIC);
			sheet.getRow(Totals_row).getCell(mapsCountCell).setCellValue(Integer.parseInt(TMapCountSS));
			sheet.getRow(Totals_row).getCell(docsCountCell).setCellType(Cell.CELL_TYPE_NUMERIC);
			sheet.getRow(Totals_row).getCell(docsCountCell).setCellValue(Integer.parseInt(TDocCountSS));

			Thread.sleep(1500);
			driver.findElement(By.xpath("//div[4]/span[2]/span/img")).click();
			driver.findElement(By.xpath("//div[@id='topNavButtons']//a[@class='includeSearch']/img")).click(); // include button

			// capturing GeoTiff total count
			Wait().until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//ul[@id='filterList']/li[2]/img[@alt='Remove Search Criteria']")));
			Wait().until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li/div/div/div/a/img")));
			String TGeoCount = driver.findElement(By.id("mapsTabLabel")).getText();
			String TGeoCountSS = TGeoCount.substring(TGeoCount.indexOf("(") + 1, TGeoCount.indexOf(")"));
			sheet.getRow(Totals_row).getCell(geoCountCell).setCellType(Cell.CELL_TYPE_NUMERIC);
			sheet.getRow(Totals_row).getCell(geoCountCell).setCellValue(Integer.parseInt(TGeoCountSS));

			// remove Geotiff filter
			Thread.sleep(1500);
			driver.findElement(By.xpath("//ul[@id='filterList']/li[2]/img[@alt='Remove Search Criteria']")).click();
			Thread.sleep(2000);
			Wait().until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='box9']/div/img[@id='capub']")));

			/*
			 * Getting the data for each publisher now -------------------------------------------
			 */

			// to expand Publisher facet section
			driver.findElement(By.xpath("//div[@id='box9']/div/img[@id='capub']")).click();
			Thread.sleep(1500);

			// capturing sorting date values for each publisher
			for (int i = 1; i <= TOTAL_PUBLISHERS; i++) {
				Wait().until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[9]/ul/li[1]/label")));
				ROW_FOR_PUBLISHER = 5 + i;
				try {
					if (driver.findElement(By.xpath("//div[@id='box9']/div[@class='vmvalinks']/a[@class='view_all']")).isDisplayed())
						driver.findElement(By.xpath("//div[@id='box9']/div[@class='vmvalinks']/a[@class='view_all']")).click();

				} catch (Exception e) {}
				Thread.sleep(2000);
				driver.findElement(By.xpath("//div[@id='box9']//label[contains(.,'" + pubNameList[i - 1] + "')]/preceding-sibling::span/span/img")).click(); // Including publisher
				driver.findElement(By.xpath("//div[@id='bottomNavButtons']/div/div/a/img")).click(); //Include
				Thread.sleep(2000);
				Wait().until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//ul[@class='display']/li[1]//a/img")));

				// Getting oldest and newest dates for each publisher
				String[] tempPubOld = driver.findElement(By.xpath("//ul[@class='display']/li[1]//font[@class='citation_d_journal']")).getText().split(",");
				String pubOld = driver.findElement(By.xpath("//ul[@class='display']/li[1]//font[@class='citation_d_journal']")).getText();

				if (tempPubOld[tempPubOld.length - 1].contains("pp."))
					pubOldest = (pubOld.substring(pubOld.lastIndexOf(",") - 4).substring(0, 4)).trim();
				else
					pubOldest = pubOld.substring(pubOld.length() - 4, pubOld.length());

				new Select(driver.findElement(By.id("sortOptions"))).selectByVisibleText("Date newest first");
				Thread.sleep(2500);
				String[] tempPubNew = driver.findElement(By.xpath("//ul[@class='display']/li[1]//font[@class='citation_d_journal']")).getText().split(",");
				String pubNew = driver.findElement(By.xpath("//ul[@class='display']/li[1]//font[@class='citation_d_journal']")).getText();

				if (tempPubNew[tempPubNew.length - 1].contains("pp."))
					pubNewest = (pubNew.substring(pubNew.lastIndexOf(",") - 4).substring(0, 4)).trim();
				else
					pubNewest = pubNew.substring(pubNew.length() - 4, pubNew.length());

				new Select(driver.findElement(By.id("sortOptions"))).selectByVisibleText("Date oldest first"); // Resetting sorting value -> Oldest
				Thread.sleep(2500);
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(OldestCell).setCellType(Cell.CELL_TYPE_NUMERIC);
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(OldestCell).setCellValue(Integer.parseInt(pubOldest));
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(NewestCell).setCellType(Cell.CELL_TYPE_NUMERIC);
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(NewestCell).setCellValue(Integer.parseInt(pubNewest));

				try {
					if (driver.findElement(By.xpath("//div[@id='box5']/div[@class='vmvalinks']/a[@class='view_all']")).isDisplayed())
						driver.findElement(By.xpath("//div[@id='box5']/div[@class='vmvalinks']/a[@class='view_all']")).click();
					Thread.sleep(2000);
				} catch (Exception e) {
					System.out.println("View All is not present for " + pubNameList[i - 1]);
					Thread.sleep(3000);
				}

				totalDecades = totalSingleYears = 0;

				totalDecades = driver.findElement(By.xpath("//div[@id='box5']")).findElements(By.xpath("//img[starts-with(@id,'pyrParent-')]")).size();
				driver.findElement(By.xpath("//img[@id='pyrParent-" + (totalDecades - 1) + "']")).click();
				totalSingleYears = driver.findElement(By.xpath("//div[@id='pyrParentBox" + (totalDecades - 1) + "']")).findElements(By.tagName("li")).size();
				String PubPYFStart = driver.findElement(By.xpath("//input[@id='pyrnav" + (totalDecades - 1) + "_" + (totalSingleYears - 1) + "']")).getAttribute("value");
				String PubPYFEnd = driver.findElement(By.xpath("//input[@id='pyrnav0_0']")).getAttribute("value");

				sheet.getRow(ROW_FOR_PUBLISHER).getCell(PYFStartCell).setCellType(Cell.CELL_TYPE_NUMERIC);
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(PYFStartCell).setCellValue(Integer.parseInt(PubPYFStart));
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(PYFEndCell).setCellType(Cell.CELL_TYPE_NUMERIC);
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(PYFEndCell).setCellValue(Integer.parseInt(PubPYFEnd));

				String PubMapCount = driver.findElement(By.id("mapsTabLabel")).getText().substring(6);
				String PubMapCountSS = PubMapCount.replaceAll("[()]", "");
				String PubDocCount = driver.findElement(By.id("documentsTabLabel")).getText().substring(10);
				String PubDocCountSS = PubDocCount.replaceAll("[()]", "");
				Thread.sleep(1500);

				driver.findElement(By.xpath("//div[4]/span[2]/span/img")).click(); // including GeoTIFF filter
				driver.findElement(By.xpath("//div[@id='topNavButtons']//a[@class='includeSearch']/img")).click();
				Thread.sleep(1500);
				Wait().until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li/div/div/div/a/img")));
				String PubGeocCount = driver.findElement(By.id("mapsTabLabel")).getText().substring(6);
				String PubGeocCountSS = PubGeocCount.replaceAll("[()]", "");

				sheet.getRow(ROW_FOR_PUBLISHER).getCell(mapsCountCell).setCellType(Cell.CELL_TYPE_NUMERIC);
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(mapsCountCell).setCellValue(Integer.parseInt(PubMapCountSS));
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(docsCountCell).setCellType(Cell.CELL_TYPE_NUMERIC);
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(docsCountCell).setCellValue(Integer.parseInt(PubDocCountSS));
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(geoCountCell).setCellType(Cell.CELL_TYPE_NUMERIC);
				sheet.getRow(ROW_FOR_PUBLISHER).getCell(geoCountCell).setCellValue(Integer.parseInt(PubGeocCountSS));

				Thread.sleep(1500);
				driver.findElement(By.xpath("//ul[@id='filterList']/li[3]/img")).click(); // excluding GeoTIFF filter
				Thread.sleep(2500);
				driver.findElement(By.xpath("//ul[@id='filterList']/li[2]/img")).click(); // excluding Publisher
				Thread.sleep(2500);

				FileOutputStream fileOut = new FileOutputStream(document);
				wb.write(fileOut);
				fileOut.close();
			}
			System.out.println("Completed\n");
		} catch (Exception e) {
			//			e.printStackTrace();
			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));
			String ex = sw.toString();
			ex = ex.substring(0, ex.indexOf("\n"));
			System.out.println(ex);
		}
	}

	@AfterTest
	public void PS_Finish() throws Exception {
		driver.quit();
		String verificationErrorString = verificationErrors.toString();
		if (!"".equals(verificationErrorString)) {
			fail(verificationErrorString);
		}
	}
}
