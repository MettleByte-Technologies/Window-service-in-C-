using System;
using System.Data;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using Oracle.ManagedDataAccess.Client;
using OfficeOpenXml;
using System.Collections.Generic;
using static OfficeOpenXml.ExcelErrorValue;

namespace MyFirstService
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
        private string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=DESKTOP-8BDLNQR)(PORT=1522)))(CONNECT_DATA=(SID=XE)));User Id=sys;Password=oracle;DBA Privilege=SYSDBA;";
        private string folderPath = @"C:\\Folder A"; // Folder containing Excel files
        private string processedFolderPath = @"C:\\Folder B"; // Folder to move processed files

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 60000; // Run every minute
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            WriteToFile("Service stopped at " + DateTime.Now);
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("Checking for new Excel files at " + DateTime.Now);
            ProcessExcelFiles();
        }

        private void ProcessExcelFiles()
        {
            try
            {
                if (!Directory.Exists(folderPath))
                {
                    WriteToFile("Folder does not exist: " + folderPath);
                    return;
                }

                string[] files = Directory.GetFiles(folderPath, "*.xlsx");
                if (files.Length == 0)
                {
                    WriteToFile("No Excel files found in folder: " + folderPath);
                    return;
                }

                foreach (string file in files)
                {
                    WriteToFile("Processing file: " + Path.GetFileName(file));

                    string orderId = Path.GetFileNameWithoutExtension(file).Split('_').Last(); // Extract order ID

                    // Process both files without conditions
                    bool isProcessed = ProcessExcelFile(file, orderId);

                    // Move file to processed folder if data was inserted
                    if (isProcessed)
                    {
                        MoveFile(file, orderId);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile("Error processing files: " + ex.Message);
            }
        }

        private bool ProcessExcelFile(string filePath, string orderId)
        {
            try
            {
                WriteToFile($"Processing file: {filePath}");

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        WriteToFile($"Worksheet not found in {filePath}");
                        return false;
                    }

                    // Make the file name check case-insensitive and more flexible
                    string tableName = null;
                    string fileName = Path.GetFileName(filePath).ToLower();  // Convert to lowercase for case-insensitive check

                    if (fileName.Contains("fa_detalle_pedido"))
                    {
                        tableName = "fa_detalle_proforma";
                    }
                    else if (fileName.Contains("fa_pedido"))
                    {
                        tableName = "fa_proforma";
                    }

                    if (tableName == null)
                    {
                        WriteToFile($"Unrecognized file type: {filePath}");
                        return false;
                    }

                    // Extract columns and data from the worksheet
                    DataTable dataTable = ExtractColumnsAndData(worksheet, tableName);

                    if (dataTable.Rows.Count == 0)
                    {
                        WriteToFile($"No matching columns found for {tableName} in {filePath}");
                        return false;
                    }

                    // Insert data into the corresponding table
                    InsertDataIntoTable(dataTable, tableName);
                    WriteToFile($"Successfully inserted data into {tableName} from {filePath}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                WriteToFile($"Error processing {filePath}: " + ex.Message);
                return false;
            }
        }


        private DataTable ExtractColumnsAndData(ExcelWorksheet worksheet, string tableName)
        {
            DataTable dt = new DataTable();
            string[] expectedColumns;

            if (tableName == "fa_detalle_proforma")
            {
                expectedColumns = new string[] { "DEPR_CODIGO_EMPRESA", "DEPR_CODIGO_BODEGA", "DEPR_CODIGO_PROFORMA", "DEPR_CODIGO_PRODUCTO", "DEPR_CANTIDAD", "DEPR_PRECIO",
                                                  "DEPR_PAGO_IVA", "DEPR_COSTO", "DEPR_CANT_DSCTO1", "DEPR_PORC_DSCTO1", "DEPR_CODIGO_DSCTO1", "DEPR_CANT_DSCTO2", "DEPR_PORC_DSCTO2",
                                                  "DEPR_CODIGO_DSCTO2", "DEPR_CANT_DSCTO3", "DEPR_PORC_DSCTO3", "DEPR_CODIGO_DSCTO3", "DEPR_CANT_DSCTO4", "DEPR_PORC_DSCTO4",
                                                  "DEPR_CODIGO_DSCTO4", "DEPR_CANT_DSCTO5", "DEPR_PORC_DSCTO5", "DEPR_CODIGO_DSCTO5", "DEPR_EXTRA", "DEPR_PRECIO_G", "DEPR_NUMERO",
                                                  "DEPR_NUMERO2", "DEPR_CARACTER", "DEPR_CARACTER2", "DEPR_VALOR_ICE","DEPR_CODIGO_PEDIDO" };

            }
            else if (tableName == "fa_proforma")
            {
                expectedColumns = new string[] { "PROF_CODIGO_EMPRESA", "PROF_CODIGO_PEDIDO", "PROF_CODIGO_CLIENTE", "PROF_CODIGO_BODEGA", "PROF_DESCUENTO_TOTAL", "PROF_TERMINAL", "PROF_TIPO", "PROF_TIPO_CLIENTE", "PROF_CODIGO_VENDEDOR", "PROF_TIPO_MONEDA",
                                                  "PROF_FECHA", "PROF_FORMA_PAGO", "PROF_CODIGO_DSCTO", "PROF_IVA", "PROF_VALOR_IVA", "PROF_ESTADO", "PROF_USUARIO", "PROF_TIPO_CAMBIO", "PROF_VALOR_DESCUENTO", "PROF_OBSERVACION", "PROF_ORDEN_COMPRA",
                                                  "PROF_COMISION", "PROF_PRECIO_VTA", "PROF_FECHA_SISTEMA", "PROF_ALFA", "PROF_TOTAL_ICE", "PROF_EXPORTADA", "PROF_ENVIADO", "PROF_FECHA_ANULACION", "PROF_NOMBRE_CLIENTE", "PROF_DIRECCION", "PROF_TELEFONO",
                                                  "PROF_CEDULA_RUC", "PROF_FECHA_ULTIMO_DESP", "PROF_FECHA_POSTERGA_VCTO", "PROF_FECHA_COBRO", "" };

            }
            else
            {
                WriteToFile($"Unrecognized table name: {tableName}");
                return null;
            }

            // Get actual columns from the first row of Excel
            Dictionary<int, string> columnMapping = new Dictionary<int, string>();
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                string columnHeader = worksheet.Cells[1, col].Text.Trim();
                if (expectedColumns.Contains(columnHeader))
                {
                    columnMapping[col] = columnHeader;
                    dt.Columns.Add(columnHeader);
                }
            }

            if (dt.Columns.Count == 0)
            {
                WriteToFile($"No matching columns found for {tableName} in {worksheet.Name}");
                return dt;
            }

            // Read data from row 2 onward
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                DataRow newRow = dt.NewRow();
                foreach (var col in columnMapping)
                {
                    newRow[col.Value] = worksheet.Cells[row, col.Key].Text.Trim();
                }
                dt.Rows.Add(newRow);
            }

            return dt;
        }


        private void InsertDataIntoTable(DataTable dataTable, string tableName)
        {
            try
            {
                using (OracleConnection conn = new OracleConnection(connectionString))
                {
                    conn.Open();
                    foreach (DataRow row in dataTable.Rows)
                    {
                        // Build the insert query dynamically
                        string query = $"INSERT INTO {tableName} ({string.Join(",", dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName))}) VALUES ({string.Join(",", dataTable.Columns.Cast<DataColumn>().Select(c => $":{c.ColumnName}"))})";

                        using (OracleCommand cmd = new OracleCommand(query, conn))
                        {
                            foreach (DataColumn col in dataTable.Columns)
                            {
                                object value = row[col];

                                // Handle date columns
                                if (col.ColumnName.Equals("PROF_FECHA", StringComparison.OrdinalIgnoreCase) ||
                                    col.ColumnName.Equals("PROF_FECHA_ENTREGA", StringComparison.OrdinalIgnoreCase) ||
                                    col.ColumnName.Equals("PROF_FECHA_SISTEMA", StringComparison.OrdinalIgnoreCase))
                                {
                                    if (value is string stringValue)
                                    {
                                        string[] dateFormats = new string[] { "yyyy-MM-dd", "dd-MM-yyyy", "MM/dd/yyyy" };
                                        if (DateTime.TryParseExact(stringValue, dateFormats, null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                                        {
                                            value = parsedDate;
                                        }
                                        else
                                        {
                                            WriteToFile($"Invalid date format in column {col.ColumnName}: {stringValue}");
                                            value = DBNull.Value;
                                        }
                                    }
                                }
                                else
                                {
                                    // Handle potential overflow for number or string columns
                                    if (value is decimal decimalValue)
                                    {
                                        // Adjust precision and scale based on your schema
                                        decimalValue = Math.Round(decimalValue, 2); // Round to 2 decimal places
                                        if (decimalValue > 99999999.99m)  // Example threshold
                                        {
                                            WriteToFile($"Value too large for column {col.ColumnName}: {decimalValue}. Truncating value.");
                                            value = 99999999.99m;  // Truncate to max allowed
                                        }
                                        else
                                        {
                                            value = decimalValue;
                                        }
                                    }
                                    else if (value is string stringValue)
                                    {
                                        // Check string length (example: 255 characters max for VARCHAR2)
                                        if (stringValue.Length > 255)
                                        {
                                            WriteToFile($"String too long for column {col.ColumnName}: {stringValue}. Truncating value.");
                                            value = stringValue.Substring(0, 255);  // Truncate string to 255 characters
                                        }
                                    }

                                    // Handle null or empty values
                                    if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                                    {
                                        value = DBNull.Value;
                                    }
                                }

                                // Add parameters for the query
                                if (value is DateTime dateValue)
                                {
                                    cmd.Parameters.Add(new OracleParameter(col.ColumnName, OracleDbType.Date)).Value = dateValue;
                                }
                                else
                                {
                                    cmd.Parameters.Add(new OracleParameter(col.ColumnName, value ?? DBNull.Value));
                                }
                            }
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile($"Error inserting data into {tableName}: " + ex.Message);
            }
        }





        private void MoveFile(string filePath, string orderId)
        {
            try
            {
                string newFileName = Path.GetFileNameWithoutExtension(filePath) + "_old.xlsx";
                string newFilePath = Path.Combine(processedFolderPath, newFileName);
                File.Move(filePath, newFilePath);
                WriteToFile($"Moved {filePath} to {newFilePath}");
            }
            catch (Exception ex)
            {
                WriteToFile($"Error moving file {filePath}: " + ex.Message);
            }
        }

        public void WriteToFile(string message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            string filepath = path + "\\ServiceLog_" + DateTime.Now.Date.ToString("yyyy-MM-dd") + ".txt";
            using (StreamWriter sw = File.AppendText(filepath))
            {
                sw.WriteLine(message);
            }
        }
    }
}
