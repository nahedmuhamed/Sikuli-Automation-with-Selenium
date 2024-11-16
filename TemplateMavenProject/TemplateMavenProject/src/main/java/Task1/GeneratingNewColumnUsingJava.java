package Task1;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.logging.Level;
import java.util.logging.Logger;

public class GeneratingNewColumnUsingJava {
    private static final Logger logger = Logger.getLogger(GeneratingNewColumnUsingJava.class.getName());
    private static final String inputFilePath = "D:\\Testing projects\\_VOISTask\\TaskData.xlsx"; // Change this to your file path
    private static final String outputFilePath = "D:\\Testing projects\\_VOISTask\\JavaUpdated.xlsx"; // Path to save the updated file

    public static void main(String[] args) {

        try (FileInputStream fileInputStream = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMMM dd, yyyy");
            LocalDate today = LocalDate.now();

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip the header row

                Cell joinDateCell = row.getCell(2); // Column C for "Join Date"
                if (joinDateCell != null && joinDateCell.getCellType() == CellType.STRING) {
                    String joinDateString = joinDateCell.getStringCellValue();

                    // Remove the weekday by splitting and taking only the month, day, and year
                    String cleanedDate = joinDateString.split(", ", 2)[1];
                    LocalDate joinDate = LocalDate.parse(cleanedDate, formatter);

                    // Calculate years in _VOIS
                    long yearsInVOIS = ChronoUnit.YEARS.between(joinDate, today);

                    // Write the result to Column D
                    Cell yearsInVOISCell = row.createCell(3); // Column D
                    yearsInVOISCell.setCellValue(yearsInVOIS);
                }
            }

            // Write the updated data back to the file
            try (FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath)) {
                workbook.write(fileOutputStream);
            }
            System.out.println("Updated Excel file saved successfully in this path " + "\"" + outputFilePath + "\" :)");

        } catch (Exception e) {
            logger.log(Level.SEVERE, "Error while generating the file", e);
        }
    }
}
