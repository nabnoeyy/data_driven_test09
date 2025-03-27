import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import  static  org.junit.jupiter.api.Assertions.assertEquals;

import java.io.File;
import java.io.FileInputStream;

public class TestRegister {
    @Test
    void test01() throws Exception {
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\hp\\Downloads\\chromedriver1\\chromedriver-win64\\chromedriver.exe");

        String path = "src/Excel09/data2.xlsx";
        FileInputStream fs = new FileInputStream(new File(path));
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int row = sheet.getLastRowNum() + 1;

        WebDriver driver = new ChromeDriver();

        for (int i = 1; i < row; i++) {
            driver.get("http://localhost/sc_shortcourses/signup");

            Row rows = sheet.getRow(i);
            if (rows == null) continue;

            // ดึงค่าจาก Excel
            String nameTitleTha = getCellValue(rows, 1);
            String firstnameTha = getCellValue(rows, 2);
            String lastnameTha = getCellValue(rows, 3);
            String nameTitleEng = getCellValue(rows, 4);
            String firstnameEng = getCellValue(rows, 5);
            String lastnameEng = getCellValue(rows, 6);
            String birthDate = getCellValue(rows, 7);
            String birthMonth = getCellValue(rows, 8);
            String birthYear = getCellValue(rows, 9);
            String idCard = getCellValue(rows, 10);
            String password = getCellValue(rows, 11);
            String mobile = getCellValue(rows, 12);
            String email = getCellValue(rows, 13);
            String address = getCellValue(rows, 14);
            String province = getCellValue(rows, 15);
            String district = getCellValue(rows, 16);
            String subDistrict = getCellValue(rows, 17);
            String postalCode = getCellValue(rows, 18);

            // กรอกข้อมูลลงฟอร์ม
            new Select(driver.findElement(By.id("nameTitleTha"))).selectByVisibleText(nameTitleTha);
            driver.findElement(By.id("firstnameTha")).sendKeys(firstnameTha);
            driver.findElement(By.id("lastnameTha")).sendKeys(lastnameTha);
            new Select(driver.findElement(By.id("nameTitleEng"))).selectByVisibleText(nameTitleEng);
            driver.findElement(By.id("firstnameEng")).sendKeys(firstnameEng);
            driver.findElement(By.id("lastnameEng")).sendKeys(lastnameEng);
            new Select(driver.findElement(By.id("birthDate"))).selectByVisibleText(birthDate);
            new Select(driver.findElement(By.id("birthMonth"))).selectByValue(birthMonth);
            new Select(driver.findElement(By.id("birthYear"))).selectByValue(birthYear);
            driver.findElement(By.id("idCard")).sendKeys(idCard);
            driver.findElement(By.id("password")).sendKeys(password);
            driver.findElement(By.id("mobile")).sendKeys(mobile);
            driver.findElement(By.id("email")).sendKeys(email);
            driver.findElement(By.id("address")).sendKeys(address);
            driver.findElement(By.id("province")).sendKeys(province);
            driver.findElement(By.id("district")).sendKeys(district);
            driver.findElement(By.id("subDistrict")).sendKeys(subDistrict);
            driver.findElement(By.id("postalCode")).sendKeys(postalCode);

            WebElement accept = driver.findElement(By.id("accept"));
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", accept);
            Thread.sleep(1000);
            if (!accept.isSelected()) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", accept);
            }

            // ใช้ JavaScriptExecutor คลิกปุ่ม Submit
            WebElement submitButton = driver.findElement(By.xpath("/html/body/section/div/div/form/div[6]/button"));
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", submitButton);

            // ตรวจสอบและคลิก checkbox 'accept'

      WebElement alertTitle = driver.findElement(By.id("swal2-title"));
      assertEquals("ลงทะเบียนสำเร็จ", alertTitle.getText());
            Thread.sleep(2000); // หน่วงเวลาให้ระบบประมวลผลก่อนไปแถวถัดไป
        }

        driver.quit();
        workbook.close();
        fs.close();
    }

    private String getCellValue(Row row, int cellIndex) {
        if (row.getCell(cellIndex) != null) {
            return row.getCell(cellIndex).toString().trim();
        }
        return "";
    }
}