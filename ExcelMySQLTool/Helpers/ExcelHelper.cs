using System;
using System.Data;
using System.IO;
using ExcelDataReader; // For .xls files
using OfficeOpenXml; // For .xlsx files (EPPlus)
using System.Text; // For CSV StringBuilder
using System.Linq; // For CSV escaping

namespace ExcelMySQLTool.Helpers
{
    public static class ExcelHelper
    {
        /// <summary>
        /// Reads an Excel file (.xls or .xlsx) into a DataTable.
        /// </summary>
        /// <param name="filePath">The path to the Excel file.</param>
        /// <param name="errorMessage">Output parameter for any error messages.</param>
        /// <returns>A DataTable containing the Excel data, or null if an error occurred.</returns>
        public static DataTable ReadExcelFile(string filePath, out string errorMessage)
        {
            errorMessage = null;
            DataTable dataTable = new DataTable();

            try
            {
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    errorMessage = "Error: File path is invalid or file does not exist.";
                    return null;
                }

                string extension = Path.GetExtension(filePath).ToLower();

                if (extension == ".xlsx")
                {
                    // EPPlus 5.x and later uses ExcelPackage.LicenseContext
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Or LicenseContext.Commercial if you have a license
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        if (package.Workbook.Worksheets.Count == 0)
                        {
                            errorMessage = "Error: The Excel file contains no worksheets.";
                            return null;
                        }

                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first sheet
                        if (worksheet.Dimension == null)
                        {
                             errorMessage = "Error: The worksheet is empty or has no dimension.";
                             return null;
                        }

                        // Add columns - assuming first row is header
                        if (worksheet.Dimension.Rows < 1)
                        {
                            errorMessage = "Error: Worksheet contains no header row.";
                            return null;
                        }
                        foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                        {
                            dataTable.Columns.Add(firstRowCell.Text);
                        }

                        // Add rows
                        for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                        {
                            var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                            DataRow row = dataTable.NewRow(); // Use NewRow and then Add to ensure it belongs to the table
                            for(int colNum = 1; colNum <= worksheet.Dimension.End.Column; colNum++)
                            {
                                // Check if column exists in DataTable (can happen if header has fewer columns than some data rows)
                                if (colNum -1 < dataTable.Columns.Count)
                                {
                                     row[colNum-1] = worksheet.Cells[rowNum, colNum].Text;
                                }
                            }
                            dataTable.Rows.Add(row);
                        }
                    }
                }
                else if (extension == ".xls")
                {
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance); // Required for ExcelDataReader on .NET Core/5+
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true // Use first row as header
                                }
                            });

                            if (result.Tables.Count > 0)
                            {
                                dataTable = result.Tables[0].Copy(); // Copy to ensure it's an independent DataTable
                            }
                            else
                            {
                                errorMessage = "Error: No tables found in the .xls file.";
                                return null;
                            }
                        }
                    }
                }
                else
                {
                    errorMessage = "Error: Unsupported file type. Please select a .xls or .xlsx file.";
                    return null;
                }

                return dataTable;
            }
            catch (Exception ex)
            {
                errorMessage = $"An error occurred while reading the Excel file: {ex.Message}";
                return null;
            }
        }

        /// <summary>
        /// Gets the top N rows from a DataTable.
        /// </summary>
        /// <param name="sourceTable">The source DataTable.</param>
        /// <param name="N">The number of rows to return.</param>
        /// <returns>A new DataTable containing the top N rows, or the original table if N is too large or table is null.</returns>
        public static DataTable GetTopNRows(DataTable sourceTable, int N)
        {
            if (sourceTable == null)
            {
                return null;
            }

            if (sourceTable.Rows.Count <= N)
            {
                return sourceTable.Copy(); // Return a copy to avoid issues if original is modified
            }

            DataTable resultTable = sourceTable.Clone(); // Clone structure
            for (int i = 0; i < N && i < sourceTable.Rows.Count; i++) // Ensure i is within bounds
            {
                resultTable.ImportRow(sourceTable.Rows[i]);
            }
            return resultTable;
        }

        /// <summary>
        /// Saves a DataTable to an Excel file (.xlsx).
        /// </summary>
        /// <param name="dataTable">The DataTable to save.</param>
        /// <param name="filePath">The path to save the Excel file.</param>
        /// <param name="errorMessage">Output parameter for any error messages.</param>
        /// <returns>True if successful, false otherwise.</returns>
        public static bool SaveDataTableToExcel(DataTable dataTable, string filePath, out string errorMessage)
        {
            errorMessage = null;
            try
            {
                if (dataTable == null)
                {
                    errorMessage = "Data to save is null.";
                    return false;
                }
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Or Commercial
                using (var package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    worksheet.Cells["A1"].LoadFromDataTable(dataTable, true); // true to print headers
                    package.SaveAs(new FileInfo(filePath));
                }
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"Error saving to Excel: {ex.Message}";
                return false;
            }
        }

        /// <summary>
        /// Saves a DataTable to a CSV file.
        /// </summary>
        /// <param name="dataTable">The DataTable to save.</param>
        /// <param name="filePath">The path to save the CSV file.</param>
        /// <param name="errorMessage">Output parameter for any error messages.</param>
        /// <returns>True if successful, false otherwise.</returns>
        public static bool SaveDataTableToCsv(DataTable dataTable, string filePath, out string errorMessage)
        {
            errorMessage = null;
            try
            {
                if (dataTable == null)
                {
                    errorMessage = "Data to save is null.";
                    return false;
                }

                StringBuilder sb = new StringBuilder();

                // Headers
                IEnumerable<string> columnNames = dataTable.Columns.Cast<DataColumn>().Select(column => QuoteCsvValue(column.ColumnName));
                sb.AppendLine(string.Join(",", columnNames));

                // Rows
                foreach (DataRow row in dataTable.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field => QuoteCsvValue(field.ToString()));
                    sb.AppendLine(string.Join(",", fields));
                }

                File.WriteAllText(filePath, sb.ToString());
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"Error saving to CSV: {ex.Message}";
                return false;
            }
        }

        private static string QuoteCsvValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            // If value contains a comma, a quote, or a newline, then quote it
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n"))
            {
                // Replace any existing quotes with double quotes
                string quotedValue = value.Replace("\"", "\"\"");
                return $"\"{quotedValue}\"";
            }
            return value;
        }
    }
}
