package selenium;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class programme {

	public static void main(String[] args) throws IOException, Throwable {

        System.setProperty("webdriver.chrome.driver", "C:\\selenium webdriver\\ChromeDriver\\chromedriver_win32\\chromedriver.exe");
        ChromeOptions opt = new ChromeOptions();
        opt.setExperimentalOption("debuggerAddress", "localhost:9988");
        WebDriver driver = new ChromeDriver(opt);

        Thread.sleep(1000);
        // Browse URL
        driver.get("https://search.google.com/search-console/removals?resource_id=sc-domain%3Apasinduljay.me");

        // ExcelSheet Path
        FileInputStream fis = new FileInputStream("E:\\Book1.xlsx");

        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        XSSFSheet sheet = workbook.getSheet("Sheet1");
        // How many rows are present
        int rowcount = sheet.getLastRowNum();

        // How many Columns are present
        int colcount = sheet.getRow(0).getLastCellNum();

        System.out.println("rowcount :" + rowcount + " colcount :" + colcount);

        for (int i = 0; i <= rowcount; i++) {
            XSSFRow celldata = sheet.getRow(i);
            String URL = celldata.getCell(0).getStringCellValue();

            Thread.sleep(500);

            // Click New request Button
            driver.findElement(By.xpath("//*[@id=\"rMvdld\"]/div/div/div/div[1]/div/div[2]/div/span/span")).click();

            Thread.sleep(500);

            // Paste URL
            driver.findElement(By.xpath("//*[@id=\"efi49d\"]/div[2]/label")).sendKeys(URL);

            System.out.println(i + "." + URL);

            Thread.sleep(1000);

            // Click Option - Remove All URLS with this prefix
            driver.findElement(By.xpath("//*[@id=\"efi49d\"]/div[3]/span/label[2]/div/div[3]/div")).click();

            Thread.sleep(1000);

            // Click Next Button
            driver.findElement(By.xpath("//*[@id=\"yDmH0d\"]/div[6]/div/div[2]/div[3]/div[2]/span/span")).click();

            Thread.sleep(1000);

            // Click Submit Button
            driver.findElement(By.xpath("//*[@id=\"yDmH0d\"]/div[6]/div/div[2]/div[3]/div[2]/span/span")).click();
            
            Thread.sleep(2300);

                     // Check if the element identified by xpath6 is present
            if (driver.findElements(By.xpath("//*[@id=\"yDmH0d\"]/div[6]/div/div[2]/div[3]/div/span")).size() > 0) 
            {
                driver.findElement(By.xpath("//*[@id=\"yDmH0d\"]/div[6]/div/div[2]/div[3]/div/span")).click();
                Thread.sleep(500);
                
            } else {
                System.out.println("Element identified by xpath6 not found. Skipping click.");
              }
        }

        // Close the browser after processing all URLs
        driver.quit();
        fis.close();
        workbook.close();
    }
}