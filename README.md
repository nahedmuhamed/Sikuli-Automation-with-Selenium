# Testing-Automation-with-Sikuli-and-Selenium
## Automation and Data Manipulation for Excel File and Web Automation for Sending Email with Selenium & Sikuli

This repository contains an automation project that integrates **Selenium** and **Sikuli** for handling Excel data and sending emails. The project involves several tasks, including reading, sorting, and manipulating data in an Excel sheet, followed by sending an email with the modified sheet using **Selenium** for browser automation. Additionally, it uses **Sikuli** for visual automation tasks like sorting data in Excel.

The project follows best practices for automation and clean code development, ensuring maintainability and scalability in future automation tasks.

## Key Features:
- **Selenium WebDriver**: Automates browser actions to interact with web services (Gmail, Outlook, etc.) for sending emails.
- **Sikuli**: Used for visually automating Excel tasks like sorting data and removing duplicates.
- **Excel Data Manipulation**: The project reads and modifies an Excel sheet to calculate the years of service for candidates based on their joining date.
- **Python**: A Python-based solution for manipulating the Excel sheet and performing the data calculations is also provided as a bonus task.

## Test Coverage:

### Task 1: Excel Data Manipulation
- **Scenario**: Read the Excel sheet and calculate the years of service for each candidate based on the "joining date" column.
- **Objective**: For each candidate, update column **D** with the number of years they have spent at the company.

### Task 2: Excel Data Sorting and Deduplication
- **Scenario**: Use Sikuli to visually automate opening the Excel sheet, sorting the data by the joining date, and removing duplicate values based on candidate names.

### Task 3: Email Automation
- **Scenario**: Use Selenium to open a web browser, log in to an email service (Gmail, Outlook, etc.), and send an email to the specified address with the modified Excel sheet as an attachment. Email credentials are stored in an application properties file for security purposes, and should be removed before submitting the code.

### Task 4: Python Data Manipulation
- **Scenario**: Repeat Task 1 using Python, demonstrating the use of Python libraries like `Pandas` for data manipulation and calculation.

## Prerequisites:

- **Java Development Kit (JDK 8+)**: Required to run the Selenium-based automation tests.
- **Maven**: For managing dependencies and running the tests.
- **Selenium WebDriver**: Used for automating the interaction with web services.
- **ChromeDriver**: used in conjunction with Selenium WebDriver for testing web applications.
- **Sikuli**: For visual automation and interacting with the Excel file.
- **Python 3.x**: For the bonus task, utilizing Python for data manipulation.
- **Pandas**: A Python library used in Task 3 for data cleaning and statistical analysis.
  
## Drive Link that have Jar files for Task1, Task2 and Task3:
- [JarFiles](https://drive.google.com/drive/folders/1mNYj1ByL_-mNaEC8RtwNo9T97PRVs-Bt?usp=drive_link)

## Conclusion:

This repository demonstrates a comprehensive automation workflow using **Selenium** and **Sikuli** to automate Excel data manipulation, sorting, and email sending. The project follows clean code principles and modular design, making it easy to extend and maintain for future use cases.


Feel free to explore the repository, suggest improvements, or contribute to its development!
