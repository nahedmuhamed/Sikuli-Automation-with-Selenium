package Task3;

import Utilities.*;
import org.sikuli.script.Screen;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Properties;
import java.util.Set;
import java.util.concurrent.TimeUnit;

public class AutomatedMailUsingSelenium {

    // Locators for the email login page and email composition fields
    private static final By EmailTextBox = By.id("i0116");
    private static final By NextButton = By.id("idSIButton9");
    private static final By PasswordTextBox = By.id("i0118");
    private static final By SignIn = By.id("idSIButton9");
    private static final By DeclineButton = By.id("declineButton");
    private static final By Menu = By.id("O365_MainLink_NavMenu");
    private static final By OutlookMail = By.id("O365_AppTile_Mail");
    private static final By NewMail = By.xpath("//span[text()='New mail']");
    private static final By ToTextBox = By.xpath("//div[@role='textbox' and @aria-label='To']");
    private static final By SubjectTextBox = By.cssSelector("input[placeholder='Add a subject'][aria-label='Add a subject']");
    private static final By BodyField = By.cssSelector("div[role='textbox'][aria-label='Message body, press Alt+F10 to exit']");
    private static final By InsertTab = By.xpath("//*[@id=\"5\"]/span");
    private static final By AttachFileTab = By.xpath("//span[text()='Attach file']");
    private static final By ChooseBrowse = By.xpath("//span[contains(text(),'Browse this computer')]");
    private static final By SendButton = By.xpath("//*[@id=\"splitButton-r10__primaryActionButton\"]");

    public static void main(String[] args) {

        // Load email and password from a properties file
        Properties props = new Properties();
        try {
            FileInputStream in = new FileInputStream("D:\\Testing projects\\_VOISTask\\TemplateMavenProject\\TemplateMavenProject\\src\\main\\resources\\app.properties");
            props.load(in);
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Get email and password values from the properties file
        String email = props.getProperty("email");
        String password = props.getProperty("password");

        // Set up ChromeDriver for Selenium
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver\\chromedriver-win64\\chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");

        WebDriver driver = new ChromeDriver(options);
        driver.get("https://login.live.com");

        try {
            // Log in to Outlook using the email and password
            Utility.sendData(driver, EmailTextBox, email);
            Utility.clickingOnElement(driver, NextButton);
            Utility.sendData(driver, PasswordTextBox, password);
            Utility.clickingOnElement(driver, SignIn);
            Utility.clickingOnElement(driver, DeclineButton);
            TimeUnit.SECONDS.sleep(10);

            // Navigate to the mail section of Outlook
            Utility.clickingOnElement(driver, Menu);
            Utility.clickingOnElement(driver, OutlookMail);

            // Switch to the new window (Outlook Mail)
            Set<String> windowHandles = driver.getWindowHandles();
            ArrayList<String> tabs = new ArrayList<>(windowHandles);
            driver.switchTo().window(tabs.get(1));

            TimeUnit.SECONDS.sleep(10);

            // Compose a new email
            Utility.clickingOnElement(driver, NewMail);
            // Fill in the recipient, subject, and body of the email
            Utility.sendData(driver, ToTextBox, "islamhasan@hotmail.com");
            Utility.sendData(driver, SubjectTextBox, "Modified Excel File");
            Utility.sendData(driver, BodyField, "Hello,\n\nPlease find the modified Excel file attached.\n\nBest regards,");

            // Attach the modified Excel file
            Utility.clickingOnElement(driver, InsertTab);
            Utility.clickingOnElement(driver, AttachFileTab);
            TimeUnit.SECONDS.sleep(2);
            Utility.clickingOnElement(driver, ChooseBrowse);
            TimeUnit.SECONDS.sleep(10);

            Screen screen = new Screen();
            screen.type("D:\\Testing projects\\_VOISTask\\TaskData.xlsx" + "\n");

            // Optional: Wait for the file upload to complete
            TimeUnit.SECONDS.sleep(10);

            // Send the email
            Utility.clickingOnElement(driver, SendButton);
            System.out.println("Email sent successfully!");

            // Clear sensitive data from the properties file
            props.setProperty("email", "");
            props.setProperty("password", "");
            FileOutputStream out = new FileOutputStream("app.properties");
            props.store(out, null);
            out.close();

        } catch (Exception e) {
            e.printStackTrace();
            driver.quit();
        }
    }
}
