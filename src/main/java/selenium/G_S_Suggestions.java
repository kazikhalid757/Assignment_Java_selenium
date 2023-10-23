package selenium;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.List;

public class GoogleSearchSuggestions {
    public static void main(String[] args) {
        // Get the current day of the week (e.g., "Thursday")
        String currentDay = getCurrentDayOfWeek();

        // Initialize WebDriver (Firefox)
        WebDriver driver = initializeFirefoxDriver();

        // Load the existing Excel file
        String excelFilePath = "/Users/kazit/Assignment_java/seleniumproject/Excel.xlsx";
        try (FileInputStream fis = new FileInputStream(new File(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Check if the current day's sheet exists in the workbook
            if (workbook.getSheetIndex(currentDay) != -1) {
                Sheet worksheet = workbook.getSheet(currentDay);

                // Process each row in the worksheet
                processWorksheetRows(driver, worksheet);

                // Save the updated Excel file
                saveWorkbookToFile(excelFilePath, workbook);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // Close the browser
            closeWebDriver(driver);
        }
    }

    private static String getCurrentDayOfWeek() {
        SimpleDateFormat dateFormat = new SimpleDateFormat("EEEE");
        return dateFormat.format(new Date());
    }

    private static WebDriver initializeFirefoxDriver() {
        return new FirefoxDriver();
    }

    private static void processWorksheetRows(WebDriver driver, Sheet worksheet) {
        for (int rowIndex = 1; rowIndex < worksheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = worksheet.getRow(rowIndex);

            Cell keywordCell = row.getCell(1);
            if (keywordCell != null) {
                String keyword = keywordCell.getStringCellValue();
                if (keyword != null && !keyword.isEmpty()) {
                    searchAndProcessKeyword(driver, row, keyword);
                }
            }
        }
    }

    private static void searchAndProcessKeyword(WebDriver driver, Row row, String keyword) {
        // ... Existing code to search and process keyword ...

        // Update the corresponding columns in the Excel file
        row.createCell(2).setCellValue(minSuggestion);
        row.createCell(3).setCellValue(maxSuggestion);
    }

    private static void saveWorkbookToFile(String excelFilePath, Workbook workbook) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
            workbook.write(fos);
        }
    }

    private static void closeWebDriver(WebDriver driver) {
        driver.quit();
    }
}
