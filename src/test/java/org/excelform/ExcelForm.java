package org.excelform;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import io.github.bonigarcia.wdm.WebDriverManager;

import java.io.*;

public class ExcelForm {
    public static void main(String[] args) throws Exception {
    	String FormUrl = "https://selenium-autofill.vercel.app/";
        String filePath = "C:\\Users\\HAI\\eclipse-workspace\\Excel\\target\\ExcelForm.xlsx";
        FileInputStream fileinput = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fileinput);
        Sheet sheet = workbook.getSheetAt(0);
        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get(FormUrl);
            driver.findElement(By.id("clearAllBtn")).click();
            Alert alert = driver.switchTo().alert();
            alert.accept();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            String fullName = row.getCell(0).toString();
            String expLevel = row.getCell(1).toString();
            String workType = row.getCell(2).toString();
            String agree    = row.getCell(3).toString();
            	Thread.sleep(200);
                driver.findElement(By.id("fullName")).clear();
                driver.findElement(By.id("fullName")).sendKeys(fullName);
                Thread.sleep(200);
                Select select = new Select(driver.findElement(By.id("experience")));
                select.selectByValue(expLevel);
                Thread.sleep(200);
                if (workType != null) {
                    if (workType.equalsIgnoreCase("Remote")) {
                        driver.findElement(By.id("radioRemote")).click();
                    } else if (workType.equalsIgnoreCase("Onsite")) {
                        driver.findElement(By.id("radioOnsite")).click();
                    }
                }
                Thread.sleep(200);
                WebElement terms = driver.findElement(By.id("termsCheckbox"));
                if (agree.equalsIgnoreCase("Yes") && !terms.isSelected()) {
                    terms.click();
                }
                Thread.sleep(200);
                driver.findElement(By.id("submitBtn")).click();
                Thread.sleep(200);
                if (row.getCell(4) == null) row.createCell(4);
                if (row.getCell(5) == null) row.createCell(5);
                try {
                    WebElement errorElement = driver.findElement(
                        By.xpath("//span[contains(@class,'error-msg') and contains(@style,'block')]")
                    );
                    String errorMsg = errorElement.getText();
                    row.getCell(4).setCellValue("FAIL");
                    row.getCell(5).setCellValue(errorMsg);
                } catch (NoSuchElementException e) {
                    row.getCell(4).setCellValue("PASS");
                    row.getCell(5).setCellValue("Form submitted successfully");
                }
            Thread.sleep(200);
            driver.navigate().refresh();
            }
        FileOutputStream fileoutput = new FileOutputStream(filePath);
        workbook.write(fileoutput);
        workbook.close();
        fileinput.close();
        fileoutput.close();
        driver.quit();
    }
}
