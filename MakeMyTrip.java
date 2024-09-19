package mmt.trip;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class MakeMyTrip {
	public static void main(String[] args) throws InterruptedException, IOException {

		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\HP\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");

		ChromeOptions options = new ChromeOptions();
		options.addArguments("--disable-blink-features=AutomationControlled");
		options.setExperimentalOption("excludeSwitches", new String[] { "enable-automation" });
//		options.addArguments(
//				"user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
		options.addArguments("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36");

		WebDriver driver = new ChromeDriver(options);
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);

		driver.get("https://www.makemytrip.com");

		// Close modal
		try {
			WebElement closeModel = wait
					.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='commonModal__close']")));
			closeModel.click();
			System.out.println("Clicked on close model");
		} catch (NoSuchElementException e) {
			System.out.println("Close model not found, proceeding.");
		}

		// Click search button
		WebElement searchButton = wait.until(ExpectedConditions
				.elementToBeClickable(By.xpath("//a[@class='primaryBtn font24 latoBold widgetSearchBtn ']")));
		searchButton.click();
		System.out.println("Clicked on search button");

		Thread.sleep(10000); // Wait for search results to load

		// Check if '200 OK' response is displayed and handle it
		try {
			WebElement responseCheck = wait
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/pre")));
			String textValue = responseCheck.getText();
			if ("200-OK".equals(textValue)) {
				driver.quit(); // Quit browser if 200-OK found
				System.out.println("Quitting due to 200-OK response");
				return;
			}
		} catch (TimeoutException e) {
			System.out.println("No 200-OK response page found, continuing.");
		}

		// Handle 'OKAY, GOT IT!' or 'Got it' buttons
		try {
			WebElement okGotItButton = driver.findElement(By.xpath("//div[@id='root']//button[.='OKAY, GOT IT!']"));
			okGotItButton.click();
			System.out.println("Clicked on OKAY, GOT IT! button");
		} catch (NoSuchElementException e) {
			try {
				WebElement gotItButton = driver.findElement(By.xpath("//button[.='Got it']"));
				gotItButton.click();
				System.out.println("Clicked on Got it button");
			} catch (NoSuchElementException ex) {
				System.out.println("No 'OKAY, GOT IT!' or 'Got it' button found, proceeding.");
			}
		}

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("FlightFares");
		Row headerRow = sheet.createRow(0);
		headerRow.createCell(0).setCellValue("Price");

		int elementsPerSection = 7;
		int totalSections = 4;
		int startIndex = 0;
		int rowIndex = 1;

		for (int section = 0; section < totalSections; section++) {
			int endIndex = startIndex + elementsPerSection - 1;

			// Process all elements in the current section
			while (true) {
				try {
					String path = "//div[@data-gslide='" + startIndex + "' and @dir='ltr']";
					WebElement pricetag = driver.findElement(By.xpath(path));
					String pr = pricetag.getText();
					pr = pr.replaceAll("[, ?  ]", "");
//                    System.out.println(pr);

					Row row = sheet.createRow(rowIndex++);
					row.createCell(0).setCellValue(pr);
					Thread.sleep(500);

					startIndex++; // Move to the next index for the next element
				} catch (NoSuchElementException e) {
					// Break the loop if no more elements are found
					break;
				}
			}

			// Move to the next section
			if (section < totalSections - 1) {
				Thread.sleep(1000);
				WebElement skip = driver.findElement(
						By.xpath("html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div[1]/div/button[2]"));
				skip.click(); // html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div[1]/div/button[2]
			}

			// Prepare for the next section
			startIndex = (section + 1) * elementsPerSection; // Set the starting index for the next section
		}

		// Save to Excel
		try (FileOutputStream fileOut = new FileOutputStream("C:\\Users\\HP\\Documents\\flightfare.xlsx")) {
			workbook.write(fileOut);
		}
		workbook.close();
		driver.quit();
		System.out.println("Browser closed and process completed.");

		// Next Step In this I read data from Previous Sheet And put down in another
		// sheet in proper format

		String inputFilePath = "C:\\Users\\HP\\Documents\\flightfare.xlsx"; // Path to the input file
		String outputFilePath = "C:\\Users\\HP\\Documents\\flightfareupdation.xlsx"; // Path for updated output

		// Step 1: Read data from the first column of the input Excel file and store it
		// in a list
		List<String> flightData = readDataFromExcel(inputFilePath);

		// Step 2: Create a new Excel sheet with three columns and process the flight
		// data
		writeDataToNewExcel(flightData, outputFilePath);
	}

	// Method to read data from the first column of the Excel file
	public static List<String> readDataFromExcel(String filePath) throws IOException {
		List<String> dataList = new ArrayList<>();
		FileInputStream fis = new FileInputStream(filePath);
		Workbook workbook = new XSSFWorkbook(fis);
		Sheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			Cell cell = row.getCell(0); // Read only the first column
			if (cell != null && cell.getCellType() == CellType.STRING) {
				String cellValue = cell.getStringCellValue().trim();
				if (!cellValue.isEmpty()) { // Skip empty cells
					dataList.add(cellValue);
				}
			}
		}
		workbook.close();
		fis.close();
		return dataList;
	}

	// Method to write processed data into a new Excel file
	public static void writeDataToNewExcel(List<String> flightData, String outputFilePath) throws IOException {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("FlightFareUpdated");

		// Create header row
		Row headerRow = sheet.createRow(0);
		headerRow.createCell(0).setCellValue("Day");
		headerRow.createCell(1).setCellValue("Date");
		headerRow.createCell(2).setCellValue("Price");

		int rowIndex = 1; // Start from row 1 since row 0 is the header

		// Step 3: Split the flight data and write it to the new Excel sheet
		for (String flight : flightData) {
			Row row = sheet.createRow(rowIndex++);

			// Check if flight string contains the expected ₹ symbol
			if (flight.contains("₹")) {
				String day = flight.substring(0, 3); // Extract Day (first 3 characters)
				String date = flight.substring(3, flight.indexOf('₹')).trim(); // Extract Date (between Day and Price)
				String price = flight.substring(flight.indexOf('₹') + 1).trim(); // Extract Price (after ₹ symbol)

				// Write the extracted data to the new Excel
				row.createCell(0).setCellValue(day); // Write Day
				row.createCell(1).setCellValue(date); // Write Date
				row.createCell(2).setCellValue(price); // Write Price
			} else {
				// Handle rows without ₹ symbol gracefully
				System.out.println("Invalid format for flight data: " + flight);
			}
		}

		// Write the workbook to the output file
		FileOutputStream fos = new FileOutputStream(outputFilePath);
		workbook.write(fos);
		fos.close();
		workbook.close();

		// Now it is working fine Now I had to compare the price and print out the 3
		// most smallest fare

		String inputFilePath1 = "C:\\Users\\HP\\Documents\\flightfareupdation.xlsx"; // Input Excel file path

		// Open the Excel file
		FileInputStream fis2 = new FileInputStream(inputFilePath1);
		Workbook workbook2 = new XSSFWorkbook(fis2);
		Sheet sheet2 = workbook2.getSheetAt(0); // Assuming the data is in the first sheet

		List<PriceEntry> priceList = new ArrayList<>();

		// Iterate over rows (skipping header)
		for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
			Row row = sheet2.getRow(i);

			if (row != null && row.getCell(2) != null) { // Ensure the price cell is not empty
				String day = row.getCell(0).getStringCellValue();
				String date = row.getCell(1).getStringCellValue();
				double price = getNumericCellValue(row.getCell(2)); // Safely get price as a double

				// Add price entry to the list
				priceList.add(new PriceEntry(day, date, price));
			}
		}

		// Sort the priceList by price in ascending order
		Collections.sort(priceList, Comparator.comparingDouble(PriceEntry::getPrice));

		// Print the three smallest prices with corresponding day and date
		System.out.println("Three Smallest Prices:");
		for (int i = 0; i < Math.min(3, priceList.size()); i++) {
			PriceEntry entry = priceList.get(i);
			System.out
					.println("Day: " + entry.getDay() + ", Date: " + entry.getDate() + ", Price: ₹" + entry.getPrice());
		}

		workbook.close();
		fis2.close();
	}

	// Helper method to safely get numeric cell values (handles numeric and string
	// cases)
	private static double getNumericCellValue(Cell cell) {
		switch (cell.getCellType()) {
		case NUMERIC:
			return cell.getNumericCellValue();
		case STRING:
			try {
				// Attempt to parse the string as a double
				return Double.parseDouble(cell.getStringCellValue().replaceAll("[^\\d.]", ""));
			} catch (NumberFormatException e) {
				// Handle the case where the string cannot be parsed into a number
				System.out.println("Error: Invalid price format at cell: " + cell.getAddress());
				return 0;
			}
		default:
			return 0;
		}
	}
}

// Helper class to store price entries
class PriceEntry {
	private String day;
	private String date;
	private double price;

	public PriceEntry(String day, String date, double price) {
		this.day = day;
		this.date = date;
		this.price = price;
	}

	public String getDay() {
		return day;
	}

	public String getDate() {
		return date;
	}

	public double getPrice() {
		return price;
	}
}
