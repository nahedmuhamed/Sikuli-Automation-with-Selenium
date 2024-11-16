package Task2;

import Utilities.Util;
import org.sikuli.script.FindFailed;
import org.sikuli.script.Key;
import org.sikuli.script.Screen;
import org.sikuli.script.Pattern;

public class DataAnalysisUsingSikuli {
    private static final String DataFilePath = "D:\\Testing projects\\_VOISTask\\JavaUpdated.xlsx" + "\n";

    public static void main(String[] args) throws FindFailed {
        Screen screen = new Screen();

        Util.minimizeAllWindows();
        // Define Patterns for Images
        Pattern ExcelIcon = Util.makePattern("Excel_Icon.png");
        Pattern FileIcon = Util.makePattern("FileTab.png");
        Pattern OpenButton = Util.makePattern("OpenButton.png");
        Pattern BrowseButton = Util.makePattern("BrowseButton.png");
        Pattern ExcelDataScreen = Util.makePattern("ExcelData.png");
        Pattern JoinDateHeader = Util.makePattern("C_Column.png");
        Pattern SortButton = Util.makePattern("SortButton.png");
        Pattern SortType = Util.makePattern("SortType.png");
        Pattern ExpendTheSelection = Util.makePattern("ExpandSelection.png");
        Pattern NameHeader = Util.makePattern("B_Column.png");
        Pattern DataTab = Util.makePattern("DataTab.png");
        Pattern DuplicatesRemoving = Util.makePattern("RemoveDup.png");
        Pattern UnselectButton = Util.makePattern("UnselectAll.png");
        Pattern NameCheckBox = Util.makePattern("NameCheck.png");

        // Refactored main script using the helper functions with fixed waits
        screen.doubleClick(ExcelIcon);
        Util.maximizeWindow();
        // Open data file
        Util.waitAndClick(screen, FileIcon);
        Util.waitAndClick(screen, OpenButton);
        Util.waitAndClick(screen, BrowseButton);
        Util.waitAndType(screen, DataFilePath);
        screen.wait(ExcelDataScreen, 10);
        // Sort data by Join Date
        Util.waitAndClick(screen, JoinDateHeader);
        Util.waitAndClick(screen, SortButton);
        Util.waitAndClick(screen, SortType);
        Util.waitAndClick(screen, ExpendTheSelection);
        screen.type(Key.ENTER);
        // Remove duplicates by Name
        Util.waitAndClick(screen, NameHeader);
        Util.waitAndClick(screen, DataTab);
        Util.waitAndClick(screen, DuplicatesRemoving);
        Util.waitAndClick(screen, ExpendTheSelection);
        screen.type(Key.ENTER);
        // Unselect all except Name and confirm removing the duplicates names
        Util.waitAndClick(screen, UnselectButton);
        screen.doubleClick(NameCheckBox);
        screen.type(Key.ENTER);
        screen.wait(3.0);
        screen.type(Key.ENTER); // Final confirmation
        System.out.println("Excel file processed successfully :)");
        //Save The changes and close Excel
        Util.saveAndClose();
    }
}
