package Utilities;

import org.sikuli.script.FindFailed;
import org.sikuli.script.Screen;
import org.sikuli.script.Pattern;

import java.awt.*;
import java.awt.event.KeyEvent;
import java.util.logging.Logger;
import java.util.logging.Level;

public class Util {
    private static final String ImagesFolderPath = "D:\\Testing projects\\_VOISTask\\Images\\";
    private static final Logger logger = Logger.getLogger(Util.class.getName());

    public static Pattern makePattern(String ImagePath) {
        return new Pattern(ImagesFolderPath + ImagePath).similar(0.2);
    }

    public static void minimizeAllWindows() {
        try {
            Robot robot = new Robot();
            Thread.sleep(1000); // Initial delay for safety

            // Press Windows + M to Minimize all Windows
            robot.keyPress(KeyEvent.VK_WINDOWS);
            robot.keyPress(KeyEvent.VK_M);
            robot.keyRelease(KeyEvent.VK_WINDOWS);
            robot.keyRelease(KeyEvent.VK_M);

        } catch (Exception e) {
            logger.log(Level.SEVERE, "Error while minimizing the windows", e);
        }
    }

    public static void maximizeWindow() {
        try {
            Robot robot = new Robot();
            Thread.sleep(25000); // wait for Excel to load completely

            // Press Alt + Space to open the system menu for the active window
            robot.keyPress(KeyEvent.VK_ALT);
            robot.keyPress(KeyEvent.VK_SPACE);
            robot.keyRelease(KeyEvent.VK_SPACE);
            robot.keyRelease(KeyEvent.VK_ALT);

            // Small delay to ensure the menu opens
            Thread.sleep(500);

            // Press 'X' to select "Maximize" from the window menu
            robot.keyPress(KeyEvent.VK_X);
            robot.keyRelease(KeyEvent.VK_X);

        } catch (Exception e) {
            logger.log(Level.SEVERE, "Error while maximizing the window", e);
        }
    }

    public static void waitAndClick(Screen screen, Pattern element) throws FindFailed {
        screen.wait(element, 25); // Fixed wait time of 25 seconds
        screen.click(element);
    }

    // Utility function to wait 10 seconds before typing text
    public static void waitAndType(Screen screen, String text) {
        try {
            Thread.sleep(10000); // Fixed wait time of 10 seconds
            screen.type(text); // Type the text
        } catch (InterruptedException e) {
            logger.log(Level.SEVERE, "Error while typing", e);
        }
    }

    public static void saveAndClose() {
        try {
            Robot robot = new Robot();
            Thread.sleep(1000); // wait for The Changes to be done

            // Press CTRL + S to Save The Changes
            robot.keyPress(KeyEvent.VK_CONTROL);
            robot.keyPress(KeyEvent.VK_S);
            robot.keyRelease(KeyEvent.VK_CONTROL);
            robot.keyRelease(KeyEvent.VK_S);

            // Small delay to the changes are done
            Thread.sleep(500);

            // Press Alt + F4 to Close The Excel
            robot.keyPress(KeyEvent.VK_ALT);
            robot.keyPress(KeyEvent.VK_F4);
            robot.keyRelease(KeyEvent.VK_ALT);
            robot.keyRelease(KeyEvent.VK_F4);

        } catch (Exception e) {
            logger.log(Level.SEVERE, "Error while saving the changes", e);
        }
    }

}
