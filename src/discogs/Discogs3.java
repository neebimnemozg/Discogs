/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package discogs;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.InputEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Random;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellAddress;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.Proxy;

/**
 *
 * @author Kirill
 */
public class Discogs3 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws InterruptedException, AWTException, FileNotFoundException, IOException {

        String genre;
        String artist;
        String master;
        String country;
        String code;
        ArrayList<String> links = new ArrayList<String>();
        String first_link = null;
        String min;
        String average;
        String max;
        String want;
        String from;
        String sold;
        int num = 0;

        String timeStamp = new SimpleDateFormat("yyyy.MM.dd HH-mm").format(Calendar.getInstance().getTime());
        System.setProperty("webdriver.chrome.driver", "C://work/1/chromedriver1.exe");
        ChromeOptions options = new ChromeOptions();
        options.setExperimentalOption("excludeSwitches", new String[]{"enable-automation"});
        ChromeDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();

        try {
            File excel = new File("my 2022.02.10.xls");
            FileInputStream fis = new FileInputStream(excel);
            HSSFWorkbook wb = new HSSFWorkbook(fis);
            HSSFSheet ws = wb.getSheetAt(0);
            HSSFSheet ws1 = wb.createSheet("offers");
            HSSFCell mastcell = null;
            HSSFCell linkcell = null;
            HSSFCell c1cell = null;
            HSSFCell c2cell = null;
            HSSFCell cardcell = null;
            HSSFCell pricecell = null;
            HSSFCell deliverycell = null;
            HSSFCell sellercell = null;
            int start = 1;
            driver.get("https://www.discogs.com/");
            Thread.sleep(5000);

            for (int i = start; i <= ws.getLastRowNum(); i++) {
                HSSFRow wr = ws.getRow(i);
                HSSFCell codecell = wr.getCell(11);
                code = codecell.toString();
                HSSFCell genrecell = wr.getCell(2);
                genre = genrecell.toString();
                HSSFCell mastercell = wr.getCell(11);
                master = mastercell.toString();
                HSSFCell countrycell = wr.getCell(5);
                country = countrycell.toString();
                HSSFCell artistcell = wr.getCell(3);
                artist = artistcell.toString();
                System.out.println(code);

                driver.get(code);

                Thread.sleep(5000);
                System.out.println(i + " / " + ws.getLastRowNum());

                List<WebElement> last = driver.findElements(By.className("item_description_title"));
                if (!last.isEmpty()) {
                    for (WebElement e1 : last) {
                        String link = e1.getAttribute("href");
                        if (null == link) {
                            link = e1.getAttribute("src");
                        }
                        if (link.contains("/item")) {
                            links.add(link);
                            System.out.println(link.toString());
                            HSSFRow wr1 = ws1.createRow(num);
                            mastcell = wr1.createCell(0);
                            mastcell.setCellValue(code);
                            linkcell = wr1.createCell(1);
                            linkcell.setCellValue(link);
                            num = num + 1;
                        }
                    }
                }

                System.out.println(links.size());
                num = num - links.size();

                List<WebElement> condition = driver.findElements(By.className("item_condition"));
                if (!condition.isEmpty()) {
                    for (WebElement e1 : condition) {
                        String link = e1.getText();
                        System.out.println(link);
                        HSSFRow wr1 = ws1.getRow(num);
                        c1cell = wr1.createCell(2);
                        c1cell.setCellValue(link);
                        num = num + 1;
                    }
                }
                num = num - links.size();
                System.out.println(num);

                List<WebElement> number = driver.findElements(By.className("item_catno"));
                if (!number.isEmpty()) {
                    for (WebElement e1 : number) {
                        String link = e1.getText();
                        System.out.println(link);
                        HSSFRow wr1 = ws1.getRow(num);
                        cardcell = wr1.createCell(3);
                        cardcell.setCellValue(link);
                        num = num + 1;
                    }
                }
                num = num - links.size();
                System.out.println(num);

                List<WebElement> price = driver.findElements(By.className("price"));
                if (!price.isEmpty()) {
                    for (WebElement e1 : price) {
                        String link = e1.getText();
                        if (link.matches(".*\\d.*")) {
                            System.out.println(link);
                            HSSFRow wr1 = ws1.getRow(num);
                            pricecell = wr1.createCell(4);
                            pricecell.setCellValue(link);
                            num = num + 1;
                        }
                    }
                }
                num = num - links.size();
                System.out.println(num);

                /*List<WebElement> delivery = driver.findElements(By.className("converted_price"));
                if (!delivery.isEmpty()) {
                    for (WebElement e1 : delivery) {
                        String link = e1.getText();
                        if (link.matches(".*\\d.*")) {
                            System.out.println(link);
                            HSSFRow wr1 = ws1.getRow(num);
                            deliverycell = wr1.createCell(6);
                            deliverycell.setCellValue(link);
                            num = num + 1;
                        }
                    }
                }
                num = num - links.size();
                System.out.println(num);*/

                List<WebElement> seller = driver.findElements(By.className("seller_info"));
                if (!seller.isEmpty()) {
                    for (WebElement e1 : seller) {
                        String link = e1.getText();
                        System.out.println(link);
                        HSSFRow wr1 = ws1.getRow(num);
                        sellercell = wr1.createCell(5);
                        sellercell.setCellValue(link);
                        num = num + 1;
                    }
                }

                System.out.println(num);

                first_link = "-";
                from = "-";
                min = "-";
                max = "-";
                average = "-";
                want = "-";
                links.clear();

                fis.close();

                FileOutputStream outputStream = new FileOutputStream(new File(timeStamp + ".xls"));
                wb.write(outputStream);
                outputStream.close();
            }
            driver.quit();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
}
