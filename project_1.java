import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.List;

public class GoogleSearchAutomation {
    public static void main(String[] args) {
        WebDriver driver = null;
        Workbook workbook = null;

        try {
            DayOfWeek today = LocalDate.now().getDayOfWeek();
            String todayName = today.toString();
            File file = new File("C:\\path\\to\\keywords.xlsx");
            FileInputStream fis = new FileInputStream(file);
            workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(todayName);

            if (sheet == null) {
                System.out.println("Sheet for today not found: " + todayName);
                return;
            }
            System.setProperty("webdriver.chrome.driver", "C:\\path\\to\\chromedriver.exe");
            driver = new ChromeDriver();
            for (Row row : sheet) {
                Cell keywordCell = row.getCell(1);
                if (keywordCell == null || keywordCell.getCellType() == CellType.BLANK) continue;

                String keyword = keywordCell.getStringCellValue();
                driver.get("https://www.google.com");
                driver.findElement(By.name("q")).sendKeys(keyword);
                driver.findElement(By.name("q")).submit();
                List<WebElement> results = driver.findElements(By.cssSelector(".LC20lb"));
                if (results.isEmpty()) {
                    row.createCell(2).setCellValue("No results found");
                    row.createCell(3).setCellValue("No results found");
                    continue;
                }

                String shortest = results.get(0).getText();
                String longest = results.get(0).getText();

                for (WebElement result : results) {
                    String text = result.getText();
                    if (text.length() < shortest.length()) shortest = text;
                    if (text.length() > longest.length()) longest = text;
                }

                row.createCell(2).setCellValue(longest);
                row.createCell(3).setCellValue(shortest);
            }
            fis.close();
            FileOutputStream fos = new FileOutputStream(file);
            workbook.write(fos);
            fos.close();

            System.out.println("Automation completed successfully!");

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) workbook.close();
                if (driver != null) driver.quit();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
