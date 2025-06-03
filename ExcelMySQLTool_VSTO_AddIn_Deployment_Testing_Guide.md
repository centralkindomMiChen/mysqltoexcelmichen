# ExcelMySQLTool VSTO Add-in: Deployment and Testing Guide (for Excel 2007)

## 1. Introduction

### Purpose
This guide provides instructions for building, deploying (installing), and testing the `ExcelMySQLToolAddIn` VSTO Add-in for Microsoft Excel 2007. It focuses on ensuring the add-in correctly launches the main `ExcelMySQLTool` Windows Forms application.

### Assumptions
*   You have successfully compiled the main `ExcelMySQLTool.exe` application and all its dependencies are available.
*   You have access to the `ExcelMySQLToolAddIn` VSTO project source code.

## 2. Prerequisites

*   **Operating System:** Windows 7 (64-bit).
*   **Microsoft Office:** Microsoft Office 2007 (with Excel 2007) installed.
*   **.NET Framework:** .NET Framework 4.5 (or the specific version targeted by the VSTO project, e.g., 4.0 if 4.5 proves problematic for Office 2007 VSTO development) must be installed.
*   **Visual Studio Tools for Office (VSTO) Runtime:** The version compatible with Office 2007 and the .NET Framework used. Typically, the "Visual Studio 2010 Tools for Office Runtime" (or a later compatible version) is required. This might be installed with Office, Visual Studio, or as a separate download from Microsoft.
*   **Visual Studio (Optional, for building):** Visual Studio (Community Edition or higher) is needed if you are building the VSTO add-in from source. Not required on the target machine if a `setup.exe` (ClickOnce) is provided.
*   **Application Files:**
    *   The `ExcelMySQLToolAddIn` project files.
    *   The compiled `ExcelMySQLTool.exe` and its dependencies (e.g., `MySql.Data.dll`, `EPPlus.dll`, `ExcelDataReader.dll`).

## 3. Building the VSTO Add-in (If not already built)

1.  **Open Solution:** Launch Visual Studio and open the `ExcelMySQLTool.sln` solution file.
2.  **Set Startup Project (Optional):** If you want to debug directly, you can right-click on the `ExcelMySQLToolAddIn` project in Solution Explorer and select "Set as StartUp Project".
3.  **Select Build Configuration:** Choose `Release` from the Solution Configurations dropdown (usually on the toolbar).
4.  **Build the Add-in:**
    *   Right-click on the `ExcelMySQLToolAddIn` project in Solution Explorer.
    *   Select **Build** (or **Rebuild**).
5.  **Locate Build Output:**
    *   After a successful build, the output files will typically be in the `ExcelMySQLToolAddIn\bin\Release\` directory.
    *   Key files include:
        *   `ExcelMySQLToolAddIn.dll`
        *   `ExcelMySQLToolAddIn.vsto`
        *   `ExcelMySQLToolAddIn.dll.manifest`
        *   Other referenced DLLs (like `Microsoft.Office.Tools.Common.v4.0.Utilities.dll`).
    *   If ClickOnce publishing is configured (see section 4.1), a `setup.exe` and related files will be in a `publish` subfolder or the location specified during publishing.

## 4. Deployment/Installation Methods for Office 2007

### 4.1. Using ClickOnce Deployment (Recommended)

ClickOnce simplifies deployment and handles registration and updates.

1.  **Publish the VSTO Add-in from Visual Studio:**
    *   In Solution Explorer, right-click the `ExcelMySQLToolAddIn` project.
    *   Select **Publish...**.
    *   The **Publish Wizard** will appear:
        *   **Specify the location to publish this application:** Choose a folder path (e.g., a network share, local folder, or web server if applicable). This is where the installation files will be created.
        *   **How will users install the application?:**
            *   "From a Web site" (if hosting on a web server).
            *   "From a UNC path or file share" (common for internal deployment).
            *   "From a CD-ROM or DVD-ROM".
        *   Follow the wizard prompts. Key settings include:
            *   **Prerequisites:** Ensure ".NET Framework 4.5" (or your target version) and "Microsoft Visual Studio Tools for Office Runtime" are checked.
            *   **Updates:** Configure how the add-in should check for updates (optional).
            *   **Options:** Review signing options. Signing the ClickOnce manifest with a certificate is recommended for trusted deployment.
    *   Click **Finish** or **Publish**. Visual Studio will build and place the deployment files in the specified publish location. This includes a `setup.exe`, the `.vsto` manifest, and application files.

2.  **Installation on Target Machine (Windows 7 with Office 2007):**
    *   Navigate to the publish location (or provide the `setup.exe` to the user).
    *   Run `setup.exe`.
    *   Follow the on-screen prompts. The installer will check for prerequisites and install the add-in.

3.  **Trusting the Add-in:**
    *   During installation or when Excel first starts after installation, the user might be prompted to trust the add-in, especially if it's not signed by an already trusted publisher.
    *   If prompted, the user should click "Install" or "Trust".
    *   Office 2007 security settings can be managed via the Trust Center (see section 4.3).

### 4.2. Manual "Development" Deployment (For testing/limited use)

This method is more complex for end-users and might require manual trust configuration. It's generally better for development testing.

1.  **Prepare Files:**
    *   Create a folder on the target Windows 7 machine (e.g., `C:\ExcelAddIns\ExcelMySQLTool`).
    *   Copy the following files from your `ExcelMySQLToolAddIn\bin\Release\` directory to this new folder:
        *   `ExcelMySQLToolAddIn.dll`
        *   `ExcelMySQLToolAddIn.dll.manifest`
        *   `ExcelMySQLToolAddIn.vsto`
        *   `Microsoft.Office.Tools.Common.v4.0.Utilities.dll`
        *   Any other DLLs specifically required by the add-in itself.
    *   Copy the main `ExcelMySQLTool.exe` and *all its dependencies* (e.g., `EPPlus.dll`, `MySql.Data.dll`, `ExcelDataReader.dll`, etc.) into the **same folder**. This ensures the add-in can find and launch the main application.

2.  **Installing via .vsto file:**
    *   On the target machine, navigate to the folder where you copied the files.
    *   Double-click the `ExcelMySQLToolAddIn.vsto` file.
    *   This should trigger the VSTO runtime to install the add-in. You may be prompted to confirm the installation and trust the publisher. Click "Install".

3.  **Registry (Advanced - Generally not for manual deployment):**
    *   Proper VSTO add-in registration involves creating specific registry keys under `HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\<YourAddInID>`.
    *   The `<YourAddInID>` is typically the assembly name or a GUID. Keys include `Manifest`, `LoadBehavior`, `FriendlyName`, and `Description`.
    *   Manually creating these is error-prone. ClickOnce or a Windows Installer (MSI) package (more complex to create) handles this automatically. For Office 2007, clicking the `.vsto` file is the simplest manual approach.

### 4.3. Security Settings in Excel 2007

If the add-in doesn't load or is disabled, check Excel's Trust Center settings:

1.  Open Microsoft Excel 2007.
2.  Click the **Office Button** (top-left).
3.  Click **Excel Options**.
4.  Select **Trust Center** from the left pane.
5.  Click the **Trust Center Settings...** button.
6.  **Add-ins:**
    *   Ensure "Require Application Add-ins to be signed by a Trusted Publisher" is unchecked if your add-in is not signed or the publisher is not yet trusted.
    *   Alternatively, if signed, add the publisher to the Trusted Publishers list.
    *   Ensure "Disable all Application Add-ins" is NOT checked.
7.  **Trusted Locations (Less common for VSTO add-ins deployed via .vsto or ClickOnce but can be a factor):**
    *   If using a manual file copy method (not recommended for distribution), you might consider adding the folder where the add-in files reside to the Trusted Locations.
8.  **Message Bar:**
    *   Ensure "Show the Message Bar in all applications when content has been blocked" is selected so you are notified if Excel blocks the add-in.

## 5. Testing the Add-in

After installing the add-in:

| ID          | Description                 | Steps                                                                                                                                                                                                  | Expected Result                                                                                                                                                              |
|-------------|-----------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| TC-VSTO-001 | Ribbon Button Visibility    | 1. Open Microsoft Excel 2007. <br>2. Look for the custom ribbon tab (e.g., "Data Tools Michen").                                                                                                           | The "Data Tools Michen" tab is visible on the Excel ribbon. Within this tab, the "Excel-MySQL Utility" group and the "Open Utility" button are present and correctly labeled. |
| TC-VSTO-002 | Launch Application          | 1. Click the "Open Utility" button on the custom ribbon tab.                                                                                                                                           | The `ExcelMySQLTool` Windows Form application launches successfully without errors.                                                                                          |
| TC-VSTO-003 | Form Interaction            | 1. With the `ExcelMySQLTool` form open, try a basic interaction, e.g., enter dummy database details and attempt to connect, or try selecting an Excel file for preview.                                | The form is responsive. Controls can be interacted with. Basic actions (even if they lead to expected errors due to dummy data) do not crash the application.                 |
| TC-VSTO-004 | Multiple Clicks on Button   | 1. With the `ExcelMySQLTool` form already open (from TC-VSTO-002), click the "Open Utility" button on the Excel ribbon again.                                                                           | A new instance of the `ExcelMySQLTool` form is NOT created. The existing, already open form is brought to the front and activated.                                        |
| TC-VSTO-005 | Excel Shutdown              | 1. Close Microsoft Excel while the `ExcelMySQLTool` form is open. <br>2. Close Microsoft Excel after closing the `ExcelMySQLTool` form.                                                                   | Excel closes cleanly without any unhandled exceptions or error messages related to the add-in.                                                                             |
| TC-VSTO-006 | Add-in Load on Startup    | 1. Close Excel completely. <br>2. Re-open Excel.                                                                                                                                                      | The "Data Tools Michen" ribbon tab and "Open Utility" button are still present, indicating the add-in loaded correctly on startup.                                       |

## 6. Troubleshooting Common Issues

*   **Add-in Not Appearing in Excel Ribbon:**
    *   **VSTO Runtime:** Ensure the "Visual Studio 2010 Tools for Office Runtime" (or a compatible version) is installed. This is a common prerequisite.
    *   **.NET Framework:** Verify that the correct version of .NET Framework (e.g., 4.5) is installed on the target machine.
    *   **Installation Issues:** If using ClickOnce, check if the installation completed successfully. Look for any error messages during setup.
    *   **Trust Center Settings:**
        *   Go to Excel Options > Trust Center > Trust Center Settings > Add-ins.
        *   Ensure "Disable all Application Add-ins" is not checked.
        *   Check "Manage: COM Add-ins" and "Manage: Disabled Items" (see below).
    *   **Registry (For advanced users):** Check `HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\ExcelMySQLToolAddIn` (or the AddInID if different). The `LoadBehavior` should typically be `3` (load at startup). If it's `0` or `2`, the add-in was disabled or not set to load.

*   **Add-in is Listed but Disabled:**
    *   In Excel, go to Office Button > Excel Options > Add-Ins.
    *   At the bottom, next to "Manage:", select "Disabled Items" from the dropdown and click "Go...".
    *   If `ExcelMySQLToolAddIn` is listed, select it and click "Enable". Restart Excel.
    *   Also check "Manage: COM Add-ins". Ensure `ExcelMySQLToolAddIn` is checked. If it shows an error under "Load Behavior", there might be an issue during startup.

*   **Error When Clicking the "Open Utility" Button:**
    *   **Main Application Missing:** Ensure `ExcelMySQLTool.exe` and all its dependencies (like `MySql.Data.dll`, `EPPlus.dll`, etc.) are correctly located. If using ClickOnce, these should be part of the deployment package. If manual deployment, they should be in the same directory as `ExcelMySQLToolAddIn.dll`.
    *   **Permissions:** The user might lack permissions to execute the application or access required resources.
    *   **Log Files:** Check the `txtLogOutput` within the `ExcelMySQLTool` application itself (if it manages to launch partially) or any system event logs for clues.

*   **Security Warnings:**
    *   If the add-in is not signed with a trusted certificate, users will see security warnings. For wider distribution, signing with a code signing certificate is recommended. For testing, users can choose to trust the add-in.

*   **Office Version Compatibility:**
    *   This guide is for Office 2007. Ensure that the VSTO project settings and any Interop Assemblies are compatible with Office 2007. Using `Embed Interop Types = True` (default for .NET 4.0+) can simplify PIA deployment.

## 7. Logging Test Results

For each test case executed, record the following information:

| Test Case ID | Description | Steps Performed | Expected Result | Actual Result | Pass/Fail | Notes (e.g., error messages, specific observations) |
|--------------|-------------|-----------------|-----------------|---------------|-----------|----------------------------------------------------|
| TC-XXX-XXX   |             |                 |                 |               |           |                                                    |

This detailed record will help track the testing progress and identify any issues that need resolution.
