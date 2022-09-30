/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package discogs;

import java.awt.AWTException;
import java.awt.Robot;
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
public class Discogs {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws InterruptedException, AWTException, FileNotFoundException, IOException {

        String genre;
        String artist;
        String title;
        String country;
        String code;
        ArrayList<String> links = new ArrayList<String>();
        String first_link = null;
        String min;
        String average;
        String want;
        String max;
        String from;
        int num = 0;

        String timeStamp = new SimpleDateFormat("yyyy.MM.dd HH-mm").format(Calendar.getInstance().getTime());
        System.setProperty("webdriver.chrome.driver", "C://work/chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
        options.setExperimentalOption("excludeSwitches", new String[]{"enable-automation"});
        ChromeDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();

        try {
            File excel = new File("rock2.xls");
            FileInputStream fis = new FileInputStream(excel);
            HSSFWorkbook wb = new HSSFWorkbook(fis);
            HSSFSheet ws = wb.getSheetAt(0);
            HSSFCell cardcell = null;
            HSSFCell fromcell = null;
            HSSFCell mincell = null;
            HSSFCell avcell = null;
            HSSFCell maxcell = null;
            HSSFCell wantcell = null;
            int start = 21988;

            for (int i = start; i < ws.getLastRowNum(); i++) {
                HSSFRow wr = ws.getRow(i);
                HSSFCell codecell = wr.getCell(7);
                code = codecell.toString();
                HSSFCell genrecell = wr.getCell(2);
                genre = genrecell.toString();
                HSSFCell titlecell = wr.getCell(4);
                title = titlecell.toString();
                HSSFCell countrycell = wr.getCell(5);
                country = countrycell.toString();
                HSSFCell artistcell = wr.getCell(3);
                artist = artistcell.toString();

                if (code.contains("/")) {
                    code = code.split(" / ")[0];
                }

                if (artist.contains(" ")) {
                    artist = artist.split(" ")[0];
                }

                if (country.contains(" ")) {
                    country = country.split(" ")[0];
                }

                if (title.contains(" ")) {
                    title = title.split(" ")[0];
                }

                driver.get("https://www.discogs.com/ru/search/?type=all&barcode=" + code + "&layout=sm");
                if (i == start) {
                    Thread.sleep(60000);
                } else {
                    Thread.sleep(10000);
                }

                System.out.println(i + " / " + ws.getLastRowNum());
                List<WebElement> body = driver.findElements(By.className("card_body"));
                GOTO:
                if (!body.isEmpty()) {
                    System.out.println(body.size());
                    for (int j = 0; j < body.size(); j++) {
                        if (body.get(j).getText().contains(country) && body.get(j).getText().contains(artist) && (body.get(j).getText().contains("LP") || body.get(j).getText().contains("12") || body.get(j).getText().contains("7"))) {
                            System.out.println(j);
                            System.out.println(body.get(j).getText());
                            num = j;
                            List<WebElement> list = body.get(j).findElements(By.xpath("//*[@href or @src]"));
                            for (WebElement e1 : list) {
                                String link = e1.getAttribute("href");
                                if (null == link) {
                                    link = e1.getAttribute("src");
                                }
                                if ((link.contains("/release/") || link.contains("/master/")) && !link.contains("/add") && !link.contains("/history")) {
                                    links.add(link);
                                }
                            }
                            break GOTO;
                        }
                    }
                }
                System.out.println(num);
                System.out.println(links.size());

                if (!links.isEmpty()) {
                    first_link = links.get(num);
                    cardcell = wr.createCell(9);
                    cardcell.setCellValue(first_link);
                    driver.get(first_link);
                    Thread.sleep(15000);
                    List<WebElement> last = driver.findElements(By.tagName("span"));
                    if (!last.isEmpty()) {
                        for (WebElement e : last) {
                            if (e.getText().contains("продаже")) {
                                System.out.println("цены " + e.getText());
                                from = e.getText();
                                fromcell = wr.createCell(13);
                                fromcell.setCellValue(from);
                            }
                        }
                    }

                    List<WebElement> price = driver.findElements(By.tagName("li"));
                    if (!price.isEmpty()) {
                        for (WebElement e : price) {
                            if (e.getText().contains("среднем")) {
                                System.out.println("price " + e.getText());
                                average = e.getText();
                                avcell = wr.createCell(10);
                                avcell.setCellValue(average);
                            }
                            if (e.getText().contains("Минимум") || e.getText().contains("меньшей")) {
                                System.out.println("price " + e.getText());
                                min = e.getText();
                                mincell = wr.createCell(11);
                                mincell.setCellValue(min);
                            }
                            if (e.getText().contains("Максимум") || e.getText().contains("большей")) {
                                System.out.println("price " + e.getText());
                                max = e.getText();
                                maxcell = wr.createCell(12);
                                maxcell.setCellValue(max);
                            }
                            if (e.getText().contains("Хотят")) {
                                System.out.println("want " + e.getText());
                                want = e.getText();
                                wantcell = wr.createCell(14);
                                wantcell.setCellValue(want);
                            }
                        }
                    }
                }
                links.clear();

                first_link = "-";
                from = "-";
                min = "-";
                max = "-";
                average = "-";

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
