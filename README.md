# Excel to MySQL Data Utility (ExcelMySQLTool)

## 1. Overview

ExcelMySQLTool is a C# Windows Forms application designed to simplify the process of transferring data between Microsoft Excel files (`.xls` and `.xlsx`) and a MySQL database. It also includes an Excel VSTO Add-in for convenient launching directly from Microsoft Excel (designed with Office 2007 compatibility in mind).

This utility is particularly useful for users who need to regularly import data from spreadsheets into a MySQL database or export data from MySQL tables into Excel or CSV formats for reporting or analysis.

**Key Features:**

*   **User-Friendly Interface:** A Windows Forms application providing clear options for database configuration, file selection, and operations.
*   **Database Configuration:** Allows users to specify MySQL connection details (host, port, database name, username, password) and the target/source table name.
*   **Excel File Handling:**
    *   Supports both `.xls` (Excel 97-2003) and `.xlsx` (Excel 2007 and later) formats.
    *   Provides a preview of the top 10 rows from the selected Excel file.
*   **MySQL Table Preview:** Allows users to preview the top 10 rows of a specified MySQL table.
*   **Data Import:** Imports data from a selected Excel sheet into a specified MySQL table.
*   **Data Export:** Exports data from a MySQL table to either `.xlsx` or `.csv` file formats.
*   **Date Range Filtering:** Enables filtering of data during export based on a specified date column and date range.
*   **Logging:** A dedicated log output panel displays status messages, errors, and operation results with timestamps.
*   **System Health Display:** A status bar shows the current system time and the application's memory usage.
*   **UTF-8 Support:** Designed to handle international characters correctly, including Chinese characters, in data, headers, and filenames (database and connection must also be UTF-8 configured).
*   **Compatibility:** Developed with consideration for Windows 7 (64-bit) and MySQL 5.4 environments.

## 2. Features Checklist

*   [x] C# WinForm for Excel to MySQL operations.
*   [x] DB name, password, table name display/modification in UI.
*   [x] Excel file selection.
*   [x] SQL execution row count and status labels (via log messages and dialogs).
*   [x] Export to `.xlsx`/`.csv` with destination choice.
*   [x] Win7 64-bit, MySQL 5.4 compatibility, UTF-8 handling.
*   [x] System time, memory usage display.
*   [ ] WinXP transparent宝石蓝 (gem blue) style (Implemented basic blue theming; exact XP-style transparency is complex and OS-dependent, not fully implemented).
*   [x] Right 1/3 for data preview and log output (achieved via SplitContainer layout).
*   [x] Direct compilation and GitHub upload (Code structure provided for user compilation; GitHub upload is user's responsibility).
*   [x] Excel plugin for Office 2007 (VSTO Add-in to launch the form).

## 3. Software and Environment Prerequisites

*   **Operating System:** Windows 7 (64-bit) or later.
*   **Microsoft .NET Framework:** Version 4.5 (or the version the application is compiled against). The ClickOnce installer for the VSTO add-in can be configured to check for and install this.
*   **Microsoft Office:** Microsoft Excel 2007 or later is required to use the VSTO Add-in. The core Windows Forms application can run without Excel installed if only dealing with `.xlsx`/`.csv` files that EPPlus/ExcelDataReader can handle independently for reading/writing.
*   **Visual Studio Tools for Office (VSTO) Runtime:** Required for the Excel Add-in to function. This is usually installed with Office or can be downloaded from Microsoft. The ClickOnce installer for the add-in can also manage this prerequisite.
*   **MySQL Server:** Access to a MySQL Server (tested with version 5.4 compatibility in mind, but should work with newer versions). Ensure the MySQL server is configured to support UTF-8 character sets for full international character support.
*   **MySQL User Account:** Credentials for a MySQL user with necessary permissions (SELECT, INSERT, CREATE TABLE if the tool is expected to create tables, etc.) on the target database.

## 4. Development Setup (For Compiling from Source)

To compile the `ExcelMySQLTool` and `ExcelMySQLToolAddIn` from source, you will need:

*   **Visual Studio:** Microsoft Visual Studio 2017 or a later version (Community Edition is free and sufficient).
*   **Required Visual Studio Workloads:**
    *   **.NET desktop development:** For the Windows Forms application (`ExcelMySQLTool`).
    *   **Office/SharePoint development:** For the VSTO Add-in (`ExcelMySQLToolAddIn`).
*   **NuGet Packages:** The solution relies on several NuGet packages for its functionality. These should be restored automatically when you open the solution in Visual Studio. If not, you can restore them manually:
    1.  Open `ExcelMySQLTool.sln` in Visual Studio.
    2.  Right-click on the Solution in Solution Explorer.
    3.  Select "Restore NuGet Packages".

    The key packages used are:
    *   `MySql.Data`: For MySQL database connectivity.
    *   `EPPlus`: For reading and writing `.xlsx` Excel files.
    *   `ExcelDataReader`: For reading older `.xls` Excel files.
    *   `ExcelDataReader.DataSet`: To easily convert Excel data to `DataSet` objects.

## 5. How to Compile

1.  **Clone or Download the Source Code:** Obtain all project files.
2.  **Open the Solution:** Launch Visual Studio and open the `ExcelMySQLTool.sln` file.
3.  **Restore NuGet Packages:** If not done automatically, right-click the solution in Solution Explorer and select "Restore NuGet Packages".

4.  **To Compile the Windows Forms Application (`ExcelMySQLTool.exe`):**
    *   In Solution Explorer, right-click on the `ExcelMySQLTool` project.
    *   Select "Set as StartUp Project".
    *   Select a build configuration (e.g., `Debug` or `Release`) from the toolbar.
    *   From the menu, choose **Build > Build ExcelMySQLTool** (or Build Solution).
    *   The compiled application (`ExcelMySQLTool.exe`) and its dependencies will be located in the `ExcelMySQLTool\bin\<Configuration>\` directory (e.g., `ExcelMySQLTool\bin\Release\`).

5.  **To Compile the Excel VSTO Add-in (`ExcelMySQLToolAddIn`):**
    *   Ensure the `ExcelMySQLTool` project has been built first, as the Add-in references it.
    *   In Solution Explorer, right-click on the `ExcelMySQLToolAddIn` project.
    *   Select "Set as StartUp Project".
    *   Select a build configuration (e.g., `Release`).
    *   From the menu, choose **Build > Build ExcelMySQLToolAddIn**.
    *   The output files (including `ExcelMySQLToolAddIn.dll`, `ExcelMySQLToolAddIn.vsto`, and `ExcelMySQLToolAddIn.dll.manifest`) will be in `ExcelMySQLToolAddIn\bin\<Configuration>\`.
    *   **For deployment:** The recommended method is using ClickOnce. Right-click on the `ExcelMySQLToolAddIn` project and select **Publish...**. Follow the wizard to create a `setup.exe` installer. Refer to the `ExcelMySQLTool_VSTO_AddIn_Deployment_Testing_Guide.md` for detailed deployment instructions.

## 6. Testing the Application and Add-in

Comprehensive testing is crucial to ensure the application and add-in function correctly, especially concerning data integrity and UTF-8 character handling.

*   **Core Windows Forms Application Testing:**
    *   Please refer to the detailed test cases outlined in the **`ExcelMySQLTool_Testing_Guide.md`** document. This guide covers UI functionality, database connections, Excel import/export (including UTF-8 scenarios), and compatibility checks.

*   **Excel VSTO Add-in Testing:**
    *   For instructions on deploying, installing, and testing the Excel VSTO Add-in (which launches the main application), please refer to the **`ExcelMySQLTool_VSTO_AddIn_Deployment_Testing_Guide.md`** document.

## 7. Delivered Files Overview

The project solution includes the following key components:

*   **`ExcelMySQLTool.sln`**: The Visual Studio Solution file containing both projects.
*   **`ExcelMySQLTool/`**: Folder containing the source code for the main Windows Forms application.
    *   `ExcelMySQLTool.csproj`: The C# project file.
    *   `Program.cs`: The main entry point for the application.
    *   `MainForm.cs`: Code-behind for the main application window.
    *   `MainForm.Designer.cs`: Designer-generated code for the main form UI.
    *   `MainForm.resx`: Resources for the main form.
    *   `Helpers/`: Subfolder containing helper classes.
        *   `ExcelHelper.cs`: Contains logic for reading from and writing to Excel files.
        *   `MySqlHelper.cs`: Contains logic for interacting with the MySQL database.
*   **`ExcelMySQLToolAddIn/`**: Folder containing the source code for the VSTO Excel Add-in.
    *   `ExcelMySQLToolAddIn.csproj`: The C# project file for the add-in.
    *   `ThisAddIn.cs`: Core class for the VSTO Add-in.
    *   `ManageDataRibbon.cs`: Code for the custom Excel ribbon tab and button.
    *   `ManageDataRibbon.Designer.cs`: Designer-generated code for the ribbon.
*   **Documentation:**
    *   `README.md` (This file): Overview, setup, compilation, and general information.
    *   `ExcelMySQLTool_Testing_Guide.md`: Detailed test cases for the main Windows Forms application.
    *   `ExcelMySQLTool_VSTO_AddIn_Guide.md`: Guide for setting up the VSTO Add-in project (development focused).
    *   `ExcelMySQLTool_VSTO_AddIn_Deployment_Testing_Guide.md`: Guide for deploying and testing the VSTO Add-in.

## 8. Notes on "WinXP Transparent Jewel Blue Style"

*   The application's main form (`MainForm`) has its `BackColor` property set to `Color.FromArgb(50, 100, 200)` to provide a "jewel blue" theme.
*   Achieving true Windows XP-style transparency and specific visual effects with standard Windows Forms controls is complex and heavily dependent on the operating system's theming capabilities. The focus of this project has been on functionality, data integrity, and compatibility.
*   Advanced custom styling beyond basic color changes would typically require third-party UI component libraries or extensive custom control painting, which is outside the scope of the initial requirements. The current styling provides a nod to the requested color theme within the constraints of standard Windows Forms development.
