using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelMySQLTool.Helpers; // Added for ExcelHelper
using System.IO; // Added for Path.GetFileName
using System.Diagnostics; // For Process for memory usage

namespace ExcelMySQLTool
{
    public partial class MainForm : Form
    {
        private DataTable _currentExcelDataTable; // To store the loaded Excel data
        private MySqlHelper _mySqlHelper;

        public MainForm()
        {
            InitializeComponent();
            // Custom initialization after InitializeComponent() call
            this.BackColor = Color.FromArgb(50, 100, 200); // Jewel blue
            // Set initial selection for ComboBox to avoid blank default
            if (cmbPreviewType.Items.Count > 0)
            {
                cmbPreviewType.SelectedIndex = 0;
            }

            // Hook up timer event
            timerSystemHealth.Tick += TimerSystemHealth_Tick;
            timerSystemHealth.Start();

            // Hook up menu item events
            exitToolStripMenuItem.Click += ExitToolStripMenuItem_Click;
            // aboutToolStripMenuItem.Click += AboutToolStripMenuItem_Click; // Placeholder for now

            // Hook up button events
            btnSelectExcelFile.Click += BtnSelectExcelFile_Click;
            btnImportToMySQL.Click += BtnImportToMySQL_Click;
            btnSelectExportLocation.Click += BtnSelectExportLocation_Click;
            btnExportFromMySQL.Click += BtnExportFromMySQL_Click;
            chkFilterByDate.CheckedChanged += ChkFilterByDate_CheckedChanged;

            // Hook up ComboBox event
            cmbPreviewType.SelectedIndexChanged += CmbPreviewType_SelectedIndexChanged;

            // Configure DataGridView
            dgvPreview.AutoGenerateColumns = true;
            dgvPreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            // Initial state for date filter controls
            ToggleDateFilterControls(false);
        }

        private void LogMessage(string message)
        {
            if (txtLogOutput.InvokeRequired)
            {
                txtLogOutput.Invoke(new Action(() => txtLogOutput.AppendText($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}")));
            }
            else
            {
                txtLogOutput.AppendText($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}");
            }
        }

        private bool InitializeMySqlHelper()
        {
            string host = txtDbHost.Text.Trim();
            string port = txtDbPort.Text.Trim();
            string dbName = txtDbName.Text.Trim();
            string user = txtDbUser.Text.Trim();
            string password = txtDbPassword.Text; // No Trim() for password

            if (string.IsNullOrEmpty(host) || string.IsNullOrEmpty(port) || string.IsNullOrEmpty(dbName) || string.IsNullOrEmpty(user))
            {
                LogMessage("Error: Database configuration is incomplete. Host, Port, DB Name, and User are required.");
                _mySqlHelper = null;
                return false;
            }

            // If _mySqlHelper is already initialized with the same parameters, no need to re-initialize.
            // This is a basic check; more sophisticated would compare all fields.
            if (_mySqlHelper != null && host == txtDbHost.Text && port == txtDbPort.Text && dbName == txtDbName.Text && user == txtDbUser.Text)
            {
                 // Test existing connection if parameters seem unchanged
                if (_mySqlHelper.TestConnection(out string quickTestError))
                {
                    // LogMessage("MySQL connection already established and tested."); // Optional: can be chatty
                    return true;
                }
                LogMessage($"Re-testing existing MySQL connection failed: {quickTestError}");
                // Proceed to re-initialize if test fails
            }


            _mySqlHelper = new MySqlHelper(host, port, dbName, user, password);
            if (!_mySqlHelper.TestConnection(out string testError))
            {
                LogMessage($"Error connecting to MySQL: {testError}");
                _mySqlHelper = null;
                return false;
            }

            LogMessage("Successfully connected to MySQL database.");
            return true;
        }


        private void BtnSelectExcelFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"; // Added xlsm
                openFileDialog.Title = "Select an Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    lblSelectedExcelFile.Text = filePath;
                    LogMessage($"Selected Excel file: {filePath}");

                    _currentExcelDataTable = ExcelHelper.ReadExcelFile(filePath, out string errorMessage);

                    if (errorMessage != null)
                    {
                        LogMessage($"Error reading Excel file: {errorMessage}");
                        dgvPreview.DataSource = null;
                        _currentExcelDataTable = null;
                    }
                    else
                    {
                        LogMessage($"Successfully read '{Path.GetFileName(filePath)}'.");
                        if (_currentExcelDataTable != null)
                        {
                            LogMessage($"Total rows read: {_currentExcelDataTable.Rows.Count}. Displaying top 10 rows in preview.");
                            dgvPreview.DataSource = ExcelHelper.GetTopNRows(_currentExcelDataTable, 10);
                            if (cmbPreviewType.SelectedItem == null || cmbPreviewType.SelectedItem.ToString() != "Excel Preview")
                            {
                                cmbPreviewType.SelectedItem = "Excel Preview";
                            }
                            else
                            {
                                RefreshPreview();
                            }
                        }
                        else
                        {
                            LogMessage("Excel file read but returned no data table (it might be empty or header only).");
                            dgvPreview.DataSource = null;
                        }
                    }
                }
            }
        }

        private void BtnImportToMySQL_Click(object sender, EventArgs e)
        {
            if (_currentExcelDataTable == null || _currentExcelDataTable.Rows.Count == 0)
            {
                LogMessage("No Excel data loaded to import, or Excel data is empty.");
                return;
            }

            string tableName = txtTableName.Text.Trim();
            if (string.IsNullOrEmpty(tableName))
            {
                LogMessage("MySQL Table Name is not specified for import.");
                MessageBox.Show("Please enter the target MySQL table name.", "Table Name Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTableName.Focus();
                return;
            }

            if (!InitializeMySqlHelper())
            {
                LogMessage("MySQL connection not initialized. Import aborted.");
                return;
            }

            LogMessage($"Starting import of {_currentExcelDataTable.Rows.Count} rows to table '{tableName}'...");
            int rowsImported = _mySqlHelper.ImportDataTable(_currentExcelDataTable, tableName, out string importMessage);

            if (rowsImported == -1) // Error occurred
            {
                LogMessage($"Import failed: {importMessage}");
                MessageBox.Show($"An error occurred during import:\n{importMessage}", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                LogMessage($"Import process completed. Successfully imported {rowsImported} rows into '{tableName}'. Additional info: {importMessage}");
                MessageBox.Show($"Successfully imported {rowsImported} rows out of {_currentExcelDataTable.Rows.Count} into '{tableName}'.\n{(string.IsNullOrEmpty(importMessage) ? "" : $"Details: {importMessage}")}",
                                "Import Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnSelectExportLocation_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|CSV (Comma delimited) (*.csv)|*.csv";
                saveFileDialog.Title = "Select Export Location and Format";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    lblSelectedExportFile.Text = saveFileDialog.FileName;
                    LogMessage($"Selected export location: {saveFileDialog.FileName}");
                }
            }
        }

        private void BtnExportFromMySQL_Click(object sender, EventArgs e)
        {
            string tableName = txtTableName.Text.Trim();
            if (string.IsNullOrEmpty(tableName))
            {
                LogMessage("MySQL Table Name is not specified for export.");
                MessageBox.Show("Please enter the MySQL table name to export from.", "Table Name Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTableName.Focus();
                return;
            }

            string exportFilePath = lblSelectedExportFile.Text;
            if (exportFilePath == "No location selected" || string.IsNullOrEmpty(exportFilePath))
            {
                LogMessage("Export location and format not selected.");
                MessageBox.Show("Please select an export location and format first.", "Export Location Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                btnSelectExportLocation.Focus();
                return;
            }

            if (!InitializeMySqlHelper())
            {
                LogMessage("MySQL connection not initialized. Export aborted.");
                return;
            }

            string dateColumn = null;
            DateTime startDate = DateTime.MinValue;
            DateTime endDate = DateTime.MaxValue;

            if (chkFilterByDate.Checked)
            {
                dateColumn = txtDateColumn.Text.Trim();
                if (string.IsNullOrEmpty(dateColumn))
                {
                    LogMessage("Date column for filtering is checked but not specified.");
                    MessageBox.Show("Please specify the date column for filtering or uncheck the filter option.", "Date Column Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtDateColumn.Focus();
                    return;
                }
                startDate = dtpStartDate.Value;
                endDate = dtpEndDate.Value;
                LogMessage($"Exporting data from '{tableName}' with date filter on column '{dateColumn}' between {startDate:yyyy-MM-dd} and {endDate:yyyy-MM-dd}.");
            }
            else
            {
                LogMessage($"Exporting all data from table '{tableName}'.");
            }

            DataTable exportedData = _mySqlHelper.ExportTable(tableName, chkFilterByDate.Checked, dateColumn, startDate, endDate, out string exportErrorMsg);

            if (exportErrorMsg != null || exportedData == null)
            {
                LogMessage($"Export failed: {exportErrorMsg ?? "No data returned."}");
                MessageBox.Show($"Export failed: {exportErrorMsg ?? "No data returned from table."}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            LogMessage($"Successfully fetched {exportedData.Rows.Count} rows from '{tableName}'. Now saving to file...");

            bool saveSuccess = false;
            string saveError = "";
            string fileExtension = Path.GetExtension(exportFilePath).ToLower();

            if (fileExtension == ".xlsx")
            {
                saveSuccess = ExcelHelper.SaveDataTableToExcel(exportedData, exportFilePath, out saveError);
            }
            else if (fileExtension == ".csv")
            {
                saveSuccess = ExcelHelper.SaveDataTableToCsv(exportedData, exportFilePath, out saveError);
            }
            else
            {
                saveError = "Unsupported file extension for export. Only .xlsx and .csv are supported.";
                LogMessage(saveError);
                MessageBox.Show(saveError, "Unsupported Format", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (saveSuccess)
            {
                LogMessage($"Data successfully exported from '{tableName}' to '{exportFilePath}'.");
                MessageBox.Show($"Data ({exportedData.Rows.Count} rows) successfully exported to:\n{exportFilePath}", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                LogMessage($"Failed to save exported data to '{exportFilePath}': {saveError}");
                MessageBox.Show($"Failed to save exported data:\n{saveError}", "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void CmbPreviewType_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshPreview();
        }

        private void RefreshPreview()
        {
            if (cmbPreviewType.SelectedItem == null) return;

            string selectedPreview = cmbPreviewType.SelectedItem.ToString();
            dgvPreview.DataSource = null; // Clear previous preview

            if (selectedPreview == "Excel Preview")
            {
                if (_currentExcelDataTable != null)
                {
                    dgvPreview.DataSource = ExcelHelper.GetTopNRows(_currentExcelDataTable, 10);
                    LogMessage("Switched to Excel Preview. Displaying top 10 rows of loaded Excel data.");
                }
                else
                {
                    LogMessage("Excel data not loaded. Cannot show Excel preview.");
                }
            }
            else if (selectedPreview == "MySQL Table Preview")
            {
                string tableName = txtTableName.Text.Trim();
                if (string.IsNullOrEmpty(tableName))
                {
                    LogMessage("MySQL Table Name is not specified for preview.");
                     // Optionally, show a message box, but log is primary for this action.
                    return;
                }

                if (!InitializeMySqlHelper())
                {
                    LogMessage("MySQL connection not initialized. Cannot preview MySQL table.");
                    return;
                }

                DataTable previewData = _mySqlHelper.GetTopNRowsFromTable(tableName, 10, out string previewErrorMsg);
                if (previewErrorMsg != null || previewData == null)
                {
                    LogMessage($"Failed to preview MySQL table '{tableName}': {previewErrorMsg ?? "No data returned."}");
                }
                else
                {
                    dgvPreview.DataSource = previewData;
                    LogMessage($"Showing top 10 rows from MySQL table '{tableName}'.");
                }
            }
        }

        private void ChkFilterByDate_CheckedChanged(object sender, EventArgs e)
        {
            ToggleDateFilterControls(chkFilterByDate.Checked);
        }

        private void ToggleDateFilterControls(bool enabled)
        {
            lblDateColumn.Enabled = enabled;
            txtDateColumn.Enabled = enabled;
            lblStartDate.Enabled = enabled;
            dtpStartDate.Enabled = enabled;
            lblEndDate.Enabled = enabled;
            dtpEndDate.Enabled = enabled;
        }

        private void TimerSystemHealth_Tick(object sender, EventArgs e)
        {
            lblStatusTime.Text = $"Time: {DateTime.Now:HH:mm:ss}";
            try
            {
                var proc = Process.GetCurrentProcess();
                lblStatusMemory.Text = $"Memory: {proc.PrivateMemorySize64 / 1024 / 1024} MB";
            }
            catch (Exception ex)
            {
                // Log minimally if this minor feature fails
                // LogMessage($"Minor error updating memory usage: {ex.Message}");
                lblStatusMemory.Text = "Memory: N/A";
            }
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        // Placeholder for future event handlers
        // private void AboutToolStripMenuItem_Click(object sender, EventArgs e) { /* Show About Box */ }
    }
}
