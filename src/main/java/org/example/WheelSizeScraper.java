package org.example;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.ThreadLocalRandom;

public class WheelSizeScraper {


// ✅ Relative to project root
    public static String  extensionPath = "extensions/uBlock0.chromium";
    private static  String EXCEL_FILE_PATH = "resources/Wheel-Size_2025_10-02.xlsx";
    private static final int RESTART_BROWSER_INTERVAL = 15;
    private static  int MAX_URLS_TO_PROCESS = 100;


    static List<String> USER_AGENTS = List.of(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/118.0",
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36",
            "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/537.36 (KHTML, like Gecko) Mobile/15E148 Safari/604.1"
    );

    private static WebDriver getChromeInstance() {
        WebDriverManager.chromedriver().setup();
        String loadExtension = "load-extension=" + extensionPath;
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--disable-gpu", "--no-sandbox",
                "--disable-blink-features=AutomationControlled", "start-maximized");
        // Set a random user-agent
        String randomUserAgent = USER_AGENTS.get(ThreadLocalRandom.current().nextInt(USER_AGENTS.size()));
        options.addArguments("user-agent=" + randomUserAgent);
        return new ChromeDriver(options);
    }


    public static void main(String[] args) {


        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(EXCEL_FILE_PATH))) {
            WebDriver driver = getChromeInstance();
            processUrls(driver, workbook);
            saveWorkbook(workbook);
            driver.quit();
        } catch (IOException e) {
            System.err.println("Error handling Excel file: " + e.getMessage());
        }
    }

    private static void processUrls(WebDriver driver, Workbook workbook) {
        Sheet urlSheet = workbook.getSheet("All Links");
        Sheet dataSheet = workbook.getSheet("Data");

        if (urlSheet == null || dataSheet == null) {
            System.err.println("Error: Sheets not found!");
            return;
        }

        int counter = 0;

        for (int i = 1; i <= urlSheet.getLastRowNum(); i++) {

            Row row = urlSheet.getRow(i);
            if (row == null) continue;

            Cell urlCell = row.getCell(0);
            if (urlCell == null) continue;

            Cell statusCell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if (("Done".equalsIgnoreCase(statusCell.getStringCellValue())) || ("No Tyres Found".equalsIgnoreCase(statusCell.getStringCellValue())))
                continue;

            String url = urlCell.getStringCellValue().trim();
            counter++;
            if (counter % RESTART_BROWSER_INTERVAL == 0) {
                driver.quit();
                driver = getChromeInstance();
            }else if(counter > MAX_URLS_TO_PROCESS){
                return;
            }
            processPage(driver, url,workbook, dataSheet, statusCell);
        }
    }

    private static void processPage(WebDriver driver, String url,Workbook workbook, Sheet dataSheet, Cell statusCell) {
        driver.get(url);
        new WebDriverWait(driver, 30).until(
                webDriver -> ((JavascriptExecutor) webDriver).executeScript("return document.readyState").equals("complete")
        );

        try {
            List<WebElement> tyreRegions = driver.findElements(By.xpath("//div[contains(@class,'trims-list ')]//div[contains(@class,'region-trim')]"));
            List<WebElement> forbiddenElements = driver.findElements(By.xpath("//h1[text()='403 Forbidden']"));
            List<WebElement> captchaElements = driver.findElements(By.xpath("//h1[contains(text(),'confirm you are human')]"));

            if (tyreRegions.isEmpty() && !forbiddenElements.isEmpty()) {
                System.out.println("Handle Forbidden Elements");
                Thread.sleep(300000);
                statusCell.setCellValue("Page not accessible");
                saveWorkbook(workbook);
                return;
            } else if(tyreRegions.isEmpty() && !captchaElements.isEmpty()) {
                System.out.println("Handle Captcha");
                Thread.sleep(30000);
                statusCell.setCellValue("Page not accessible");
                saveWorkbook(workbook);
                return;
            }else if(tyreRegions.isEmpty()) {
                System.out.println("No Tyres Found");
                statusCell.setCellValue("No Tyres Found");
                saveWorkbook(workbook);
                return;
            }

            for (WebElement tyresRegion : tyreRegions) {
                extractAndSaveTyreData(statusCell, tyresRegion, url, dataSheet);
            }

        } catch (NoSuchElementException e) {
            System.err.println("Element not found on page: " + url);
            statusCell.setCellValue("Failed");
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
    }

    private static void extractAndSaveTyreData(Cell statusCell,WebElement tyresRegion, String url, Sheet dataSheet) {
        String pattern = getElementText(tyresRegion, ".//h4[contains(@class, 'panel-hdr-text')]//span[contains(@class, 'hidden-sm-down')]");
        String yearModel = getElementAttribute(tyresRegion, ".//h4/span[2]", "data-trim-name");
        String description = getElementText(tyresRegion, ".//div[contains(@class,'row data-parameters')]");

        List<WebElement> tyres = tyresRegion.findElements(By.xpath(".//tr[contains(@class,'stock') or contains(@class,'aftermarket')]"));

        int lastDataRow = dataSheet.getLastRowNum()+1;


        for (WebElement tyre : tyres) {
            Row dataRow = dataSheet.createRow(lastDataRow++);
            dataRow.createCell(0).setCellValue(url);
            dataRow.createCell(1).setCellValue(getElementText(tyre, ".//td[contains(@class,'data-tire')]"));
            dataRow.createCell(2).setCellValue(getElementAttribute(tyre, ".//td[contains(@class,'data-rim')]", "textContent"));
            dataRow.createCell(3).setCellValue(getElementAttribute(tyre, ".//td[contains(@class, 'data-offset-range')]", "textContent"));
            dataRow.createCell(4).setCellValue(getElementAttribute(tyre, ".//td[contains(@class, 'data-backspacing')]", "textContent"));
            dataRow.createCell(5).setCellValue(getElementAttribute(tyre, ".//td[contains(@class, 'data-weight')]", "textContent"));
            dataRow.createCell(6).setCellValue(getElementAttribute(tyre, ".//td[contains(@class, 'data-pressure')]", "textContent"));
            dataRow.createCell(7).setCellValue(pattern);
            dataRow.createCell(8).setCellValue(yearModel);
            dataRow.createCell(9).setCellValue(description);
        }

        if(!tyres.isEmpty()) statusCell.setCellValue("Done") ;

        System.out.println("Extracted & Updated: " + url);
    }

    private static String getElementText(WebElement element, String xpath) {
        try {
            return element.findElement(By.xpath(xpath)).getText().trim();
        } catch (NoSuchElementException e) {
            return "N/A";
        }
    }

    private static String getElementAttribute(WebElement element, String xpath, String attribute) {
        try {
            return element.findElement(By.xpath(xpath)).getAttribute(attribute).trim();
        } catch (NoSuchElementException e) {
            return "N/A";
        }
    }

    private static void saveWorkbook(Workbook workbook) {
        try (FileOutputStream fos = new FileOutputStream(EXCEL_FILE_PATH)) {
            workbook.write(fos);
            System.out.println("Excel file updated successfully!");
        } catch (IOException e) {
            System.err.println("Error saving Excel file: " + e.getMessage());
        }
    }
}
