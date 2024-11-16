package Utilities;

import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;


public class Utility {
    private Utility()
    {
        super();
    }

    public static void clickingOnElement(WebDriver driver, By locator) {
        new WebDriverWait(driver, 40) // seconds as int for older versions
                .until(ExpectedConditions.elementToBeClickable(locator)).click();
    }

    // Sending data to an element
    public static void sendData(WebDriver driver, By locator, String data) {
        new WebDriverWait(driver, 40) // seconds as int for older versions
                .until(ExpectedConditions.visibilityOfElementLocated(locator)).sendKeys(data);
    }

    // Get text from an element
    public static String getText(WebDriver driver, By locator) {
        return new WebDriverWait(driver, 40) // seconds as int for older versions
                .until(ExpectedConditions.visibilityOfElementLocated(locator)).getText();
    }

    // General wait method (works without Duration)
    public static WebDriverWait generalWait(WebDriver driver) {
        return new WebDriverWait(driver, 40); // seconds as int for older versions
    }

    // Scrolling to an element
    public static void scrolling(WebDriver driver, By locator) {
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", findWebElement(driver, locator));
    }

    // Selecting an option from a dropdown
    public static void selectingFromDropdown(WebDriver driver, By locator, String option) {
        new Select(findWebElement(driver, locator)).selectByVisibleText(option);
    }

    // Find WebElement from By locator
    public static WebElement findWebElement(WebDriver driver, By locator) {
        return driver.findElement(locator);
    }
}
