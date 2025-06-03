# ExcelMySQLTool VSTO Add-in Setup Guide

This guide outlines the steps to create a VSTO Add-in for Microsoft Excel that will launch the ExcelMySQLTool Windows Form application.

## 1. VSTO Project Setup

Follow these steps to create the VSTO Add-in project within your existing `ExcelMySQLTool.sln` solution:

1.  **Open your Solution:** Open the `ExcelMySQLTool.sln` solution in Visual Studio.
2.  **Add New Project:**
    *   In Solution Explorer, right-click on the Solution node (e.g., `Solution 'ExcelMySQLTool' (2 of 2 projects)` if you already have the WinForms project).
    *   Select **Add > New Project...**.
3.  **Choose Project Template:**
    *   In the "Add a new project" dialog, search for "Excel VSTO Add-in".
    *   Select the "Excel VSTO Add-in" template. Depending on your Visual Studio version, this might be listed under "Office/SharePoint", "Visual C# > Office", or similar categories.
    *   Click **Next**.
4.  **Configure Project:**
    *   **Project name:** Enter `ExcelMySQLToolAddIn`.
    *   **Location:** Ensure it's within your main solution directory or a suitable sub-directory.
    *   **Solution:** Make sure "Add to solution" is selected if you created it from the solution context, or "Create new solution" if you are starting fresh (though adding to the existing solution is recommended here).
    *   **Framework:** Select **.NET Framework 4.5**. If compatibility issues with Office 2007 arise during development or deployment, you might need to consider .NET Framework 4.0, but start with 4.5 as it aligns with the main application.
    *   Click **Create**.
5.  **Choose an Office application for your add-in:**
    *   A dialog titled "Create a New Document Based on an Office Application" or similar might appear (this is more common for document-level add-ins, but VSTO add-ins might also prompt for a host application confirmation).
    *   Ensure "Microsoft Excel" is selected or implied by the project template.
    *   Visual Studio will generate the basic VSTO Add-in project structure, including a `ThisAddIn.cs` file.

## 2. Reference Main Application Project

To allow the Add-in to launch the Windows Form from the `ExcelMySQLTool` project, you need to add a project reference:

1.  In **Solution Explorer**, expand the `ExcelMySQLToolAddIn` project.
2.  Right-click on **References**.
3.  Select **Add Reference...**.
4.  In the Reference Manager dialog:
    *   Select the **Projects** tab on the left.
    *   Check the box next to `ExcelMySQLTool`.
    *   Click **OK**.

## 3. `ThisAddIn.cs` Modifications

The `ThisAddIn.cs` file contains the startup and shutdown logic for your add-in. For this specific task of launching a form via a ribbon button, minimal changes are needed here.

```csharp
// ThisAddIn.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelMySQLToolAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // No specific code needed here for now if using Ribbon Designer
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Clean up resources if any were allocated by the add-in
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
```

## 4. Create Ribbon Button (Using Ribbon Designer)

1.  **Add New Ribbon Item:**
    *   In **Solution Explorer**, right-click on the `ExcelMySQLToolAddIn` project.
    *   Select **Add > New Item...**.
    *   In the "Add New Item" dialog, select **Visual C# Items > Office > Ribbon (Visual Designer)**.
    *   Name the file `ManageDataRibbon.cs`.
    *   Click **Add**. This will open the Ribbon Designer.

2.  **Design the Ribbon:**
    *   **Add a Tab:**
        *   If not already present, drag a **Tab** control from the **Toolbox** (under "Office Ribbon Controls") onto the Ribbon designer surface.
        *   Select the Tab. In the **Properties** window:
            *   Set the `Label` property to `Data Tools Michen` (or your preferred name).
            *   (Optional) Set the `ControlId` -> `ControlIdType` to `Custom`.
            *   (Optional) Set the `ControlId` -> `OfficeId` if you want to use a built-in tab, or leave as custom.
    *   **Add a Group:**
        *   Drag a **Group** control from the Toolbox onto the Tab you just created/selected.
        *   Select the Group. In the **Properties** window:
            *   Set the `Label` property to `Excel-MySQL Utility`.
    *   **Add a Button:**
        *   Drag a **Button** control from the Toolbox into the Group you just added.
        *   Select the Button. In the **Properties** window:
            *   Set the `Label` property to `Open Utility`.
            *   Set the `ControlSize` property to `RibbonControlSizeLarge`.
            *   (Optional) Set the `OfficeImageId` property to an icon name (e.g., `DatabaseSqlServer`, `TableShare`, `DataRefresh`). You can find a list of OfficeImageIds online.
            *   (Optional) Set the `ScreenTip` to "Open the Excel-MySQL Data Utility".
            *   (Optional) Set the `SuperTip` to "Launches the Excel to MySQL Data Utility tool for importing and exporting data."

3.  **Create Event Handler:**
    *   **Double-click** the "Open Utility" button you just added in the Ribbon Designer. This will automatically create the `buttonOpenUtility_Click` event handler in the `ManageDataRibbon.cs` code-behind file and switch you to the code view.

## 5. Implement Button Click Handler (`ManageDataRibbon.cs`)

Paste the following code into the generated `buttonOpenUtility_Click` event handler within `ManageDataRibbon.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ExcelMySQLTool; // Add this using statement to reference your Windows Forms project

namespace ExcelMySQLToolAddIn
{
    public partial class ManageDataRibbon
    {
        private void ManageDataRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonOpenUtility_Click(object sender, RibbonControlEventArgs e)
        {
            // Check if the form is already open to prevent multiple instances
            // This is a simple check; more robust solutions might involve a static reference
            // or checking Application.OpenForms by a unique name if you set one.
            foreach (System.Windows.Forms.Form openForm in System.Windows.Forms.Application.OpenForms)
            {
                if (openForm.GetType() == typeof(ExcelMySQLTool.MainForm))
                {
                    openForm.Activate(); // Bring to front if already open
                    return;
                }
            }

            // If not open, create and show a new instance
            ExcelMySQLTool.MainForm mainForm = new ExcelMySQLTool.MainForm();
            mainForm.Show(); // Use Show() for a non-modal window, allowing interaction with Excel.
                             // Use ShowDialog() if you want the form to be modal, blocking Excel interaction until closed.
        }
    }
}
```

**Note:**
*   Ensure the namespace `ExcelMySQLTool` matches the namespace of your `MainForm` in the Windows Forms project. If it's different, adjust the `using` statement and the type check accordingly.
*   The check for existing open forms is basic. For more complex scenarios, managing the form instance (e.g., via a static property in `ThisAddIn.cs` or a singleton pattern for the form) might be considered, but for simple cases, this is often sufficient.

## 6. Office 2007 Compatibility Considerations

*   **.NET Framework Version:** Office 2007 VSTO add-ins primarily support .NET Framework 3.5. While .NET Framework 4.x can be targeted, it requires the .NET Framework 4.x (or higher, up to 4.8 typically) to be installed on the end-user's machine. Office 2007 itself does not come with .NET 4.x.
    *   If strictly targeting Office 2007 without ensuring .NET 4.5 is on all client machines, .NET Framework 3.5 would be a safer bet for the VSTO add-in project, but this might create compatibility issues with the `ExcelMySQLTool` project if it uses .NET 4.5 features or libraries.
    *   The current plan targets .NET 4.5 for `ExcelMySQLTool`. If this is a hard requirement, then target machines for the add-in must also have .NET 4.5.
*   **VSTO Runtime:** Users will need the Visual Studio 2010 Tools for Office Runtime (or a later compatible version that supports Office 2007 add-ins built with newer Visual Studio versions). This is often distributed as a prerequisite.
*   **Primary Interop Assemblies (PIAs):** Ensure the Office 2007 Primary Interop Assemblies are available during development and potentially for deployment if not using embedding interop types (which is default for .NET 4.0+ projects). Visual Studio usually handles this, but it's a point to be aware of for older Office versions.
*   **Ribbon UI:** The Ribbon (Visual Designer) is compatible with Office 2007.
*   **Deployment:** Consider using ClickOnce deployment, which can help manage prerequisites like the .NET Framework and VSTO Runtime.

**Recommendation:**
If strict Office 2007 compatibility without additional .NET framework installations is paramount, both the WinForms application and the VSTO add-in might need to target an older .NET Framework version (e.g., 3.5 or 4.0 Client Profile). However, since the plan specifies .NET 4.5 for the main tool, assume that target machines will have .NET 4.5 available. The provided instructions target .NET 4.5 for the VSTO add-in as well for consistency. Test thoroughly on a clean Windows 7 + Office 2007 environment.
```
