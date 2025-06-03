using System;
using System.Collections.Generic;
using System.Data;
using MySql.Data.MySqlClient;
using System.Text;

namespace ExcelMySQLTool.Helpers
{
    public class MySqlHelper
    {
        private readonly string _connectionString;
        private readonly string _host;
        private readonly string _port;
        private readonly string _database;
        private readonly string _user;
        private readonly string _password;

        public MySqlHelper(string host, string port, string database, string user, string password)
        {
            _host = host;
            _port = port;
            _database = database;
            _user = user;
            _password = password;

            // Construct the connection string
            // Added AllowUserVariables=True for potential advanced scenarios, though not strictly needed for these methods.
            // Added TreatTinyAsBoolean=false for more consistent data type handling if TINYINT(1) is used for non-boolean purposes.
            _connectionString = $"Server={_host};Port={_port};Database={_database};Uid={_user};Pwd={_password};Charset=utf8;AllowUserVariables=True;TreatTinyAsBoolean=false;";
        }

        private string QuoteIdentifier(string identifier)
        {
            return $"`{identifier.Replace("`", "``")}`";
        }

        public bool TestConnection(out string errorMessage)
        {
            errorMessage = null;
            try
            {
                using (var connection = new MySqlConnection(_connectionString))
                {
                    connection.Open();
                    connection.Close();
                }
                return true;
            }
            catch (MySqlException ex)
            {
                errorMessage = $"MySQL Connection Error: {ex.Message} (Error Code: {ex.Number})";
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = $"General Connection Error: {ex.Message}";
                return false;
            }
        }

        public int ImportDataTable(DataTable dataTable, string tableName, out string errorMessage)
        {
            errorMessage = null;
            int rowsImported = 0;

            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                errorMessage = "No data to import.";
                return 0; // Not -1 as per instruction, 0 rows imported.
            }

            if (string.IsNullOrWhiteSpace(tableName))
            {
                errorMessage = "Table name not specified.";
                return -1;
            }

            StringBuilder sqlBuilder = new StringBuilder();
            sqlBuilder.Append($"INSERT INTO {QuoteIdentifier(tableName)} (");

            List<string> columnNames = new List<string>();
            foreach (DataColumn column in dataTable.Columns)
            {
                columnNames.Add(QuoteIdentifier(column.ColumnName));
            }
            sqlBuilder.Append(string.Join(", ", columnNames));
            sqlBuilder.Append(") VALUES (");

            List<string> parameterPlaceholders = new List<string>();
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                parameterPlaceholders.Add($"@param{i}");
            }
            sqlBuilder.Append(string.Join(", ", parameterPlaceholders));
            sqlBuilder.Append(");");

            string sql = sqlBuilder.ToString();

            using (var connection = new MySqlConnection(_connectionString))
            {
                MySqlTransaction transaction = null;
                try
                {
                    connection.Open();
                    transaction = connection.BeginTransaction();

                    using (var command = new MySqlCommand(sql, connection, transaction))
                    {
                        foreach (DataRow row in dataTable.Rows)
                        {
                            command.Parameters.Clear();
                            for (int i = 0; i < dataTable.Columns.Count; i++)
                            {
                                object value = row[i];
                                if (value == DBNull.Value)
                                {
                                    command.Parameters.AddWithValue($"@param{i}", null);
                                }
                                else
                                {
                                    command.Parameters.AddWithValue($"@param{i}", value);
                                }
                            }
                            int result = command.ExecuteNonQuery();
                            if (result > 0)
                            {
                                rowsImported++;
                            }
                        }
                    }
                    transaction.Commit();
                    return rowsImported;
                }
                catch (MySqlException ex)
                {
                    errorMessage = $"MySQL Import Error: {ex.Message} (Error Code: {ex.Number})";
                    try { transaction?.Rollback(); } catch { /* ignored */ }
                    return -1;
                }
                catch (Exception ex)
                {
                    errorMessage = $"General Import Error: {ex.Message}";
                    try { transaction?.Rollback(); } catch { /* ignored */ }
                    return -1;
                }
            }
        }

        public DataTable ExportTable(string tableName, bool filterByDate, string dateColumn, DateTime startDate, DateTime endDate, out string errorMessage)
        {
            errorMessage = null;
            DataTable dataTable = new DataTable();

            if (string.IsNullOrWhiteSpace(tableName))
            {
                errorMessage = "Table name not specified for export.";
                return null;
            }

            StringBuilder sqlBuilder = new StringBuilder();
            sqlBuilder.Append($"SELECT * FROM {QuoteIdentifier(tableName)}");

            List<MySqlParameter> parameters = new List<MySqlParameter>();

            if (filterByDate)
            {
                if (string.IsNullOrWhiteSpace(dateColumn))
                {
                    errorMessage = "Date column for filtering not specified.";
                    return null;
                }
                sqlBuilder.Append($" WHERE {QuoteIdentifier(dateColumn)} BETWEEN @StartDate AND @EndDate");
                parameters.Add(new MySqlParameter("@StartDate", MySqlDbType.DateTime) { Value = startDate });
                parameters.Add(new MySqlParameter("@EndDate", MySqlDbType.DateTime) { Value = endDate });
            }

            sqlBuilder.Append(";");

            try
            {
                using (var connection = new MySqlConnection(_connectionString))
                {
                    connection.Open();
                    using (var command = new MySqlCommand(sqlBuilder.ToString(), connection))
                    {
                        if (parameters.Count > 0)
                        {
                            command.Parameters.AddRange(parameters.ToArray());
                        }

                        using (var adapter = new MySqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                    }
                }
                return dataTable;
            }
            catch (MySqlException ex)
            {
                errorMessage = $"MySQL Export Error: {ex.Message} (Error Code: {ex.Number})";
                return null;
            }
            catch (Exception ex)
            {
                errorMessage = $"General Export Error: {ex.Message}";
                return null;
            }
        }

        public DataTable GetTopNRowsFromTable(string tableName, int N, out string errorMessage)
        {
            errorMessage = null;
            DataTable dataTable = new DataTable();

            if (string.IsNullOrWhiteSpace(tableName))
            {
                errorMessage = "Table name not specified for preview.";
                return null;
            }

            if (N <= 0)
            {
                errorMessage = "Number of rows (N) must be positive.";
                return null;
            }

            // LIMIT clause parameterization is a bit tricky with MySqlConnector/MySQL.Data.
            // The value for LIMIT must be an integer literal or a user variable for prepared statements.
            // Standard parameter binding @N might not work directly in LIMIT for all provider versions or configurations.
            // Using string concatenation for N is generally safe if N is strictly an int.
            // Alternatively, use user variables if the driver supports it well in this context.
            // For simplicity and common compatibility, direct integer injection for LIMIT is often used,
            // with the C# type system ensuring N is an integer.
            string sql = $"SELECT * FROM {QuoteIdentifier(tableName)} LIMIT {N};";


            try
            {
                using (var connection = new MySqlConnection(_connectionString))
                {
                    connection.Open();
                    using (var command = new MySqlCommand(sql, connection))
                    {
                        // command.Parameters.AddWithValue("@N", N); // This might not work for LIMIT depending on provider/version
                        using (var adapter = new MySqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                    }
                }
                return dataTable;
            }
            catch (MySqlException ex)
            {
                errorMessage = $"MySQL Preview Error: {ex.Message} (Error Code: {ex.Number})";
                return null;
            }
            catch (Exception ex)
            {
                errorMessage = $"General Preview Error: {ex.Message}";
                return null;
            }
        }
    }
}
