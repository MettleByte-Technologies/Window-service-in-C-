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
            timer.Interval = 6000; // Run every minute
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
                        tableName = "fa_detalle_pedido";
                    }
                    else if (fileName.Contains("fa_pedido"))
                    {
                        tableName = "fa_pedido";
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

            if (tableName == "fa_detalle_pedido")
            {
                expectedColumns = new string[] {
                                                "DEPE_CODIGO_EMPRESA", "DEPE_CODIGO_BODEGA", "DEPE_CODIGO_PEDIDO", "DEPE_CODIGO_PRODUCTO", "DEPE_CANTIDAD", "DEPE_PRECIO",
                                                "DEPE_PAGO_IVA", "DEPE_COSTO", "DEPE_CANT_DSCTO1", "DEPE_PORC_DSCTO1", "DEPE_CODIGO_DSCTO1", "DEPE_CANT_DSCTO2", "DEPE_PORC_DSCTO2",
                                                "DEPE_CODIGO_DSCTO2", "DEPE_CANT_DSCTO3", "DEPE_PORC_DSCTO3", "DEPE_CODIGO_DSCTO3", "DEPE_CANT_DSCTO4", "DEPE_PORC_DSCTO4",
                                                "DEPE_CODIGO_DSCTO4", "DEPE_CANT_DSCTO5", "DEPE_PORC_DSCTO5", "DEPE_CODIGO_DSCTO5", "DEPE_FECHA_ENTREGA", "DEPE_PRECIO_LISTA",
                                                "DEPE_CANTIDAD_PEDIDO", "DEPE_CANTIDAD_OBS", "DEPE_EXTRA", "DEPE_PRECIO_G", "DEPE_NUMERO", "DEPE_NUMERO2", "DEPE_CARACTER",
                                                "DEPE_CARACTER2", "DEPE_BACKORDER", "DEPE_ENVIO_MAIL", "DEPE_VALOR_ICE"
                                            };


            }
            else if (tableName == "fa_pedido")
            {
                expectedColumns = new string[] {
                                                    "PEDI_CODIGO_EMPRESA", "PEDI_TIPO", "PEDI_TIPO_CLIENTE", "PEDI_CODIGO_PEDIDO", "PEDI_ORDEN_COMPRA", "PEDI_CODIGO_CLIENTE",
                                                    "PEDI_NOMBRE_CLIENTE", "PEDI_DIRECCION", "PEDI_TELEFONO", "PEDI_CEDULA_RUC", "PEDI_CODIGO_BODEGA", "PEDI_CODIGO_VENDEDOR",
                                                    "PEDI_COMISION", "PEDI_PRECIO_VTA", "PEDI_FECHA", "PEDI_FECHA_ENTREGA", "PEDI_FECHA_ULTIMO_DESP", "PEDI_TIPO_MONEDA",
                                                    "PEDI_TIPO_CAMBIO", "PEDI_VALOR_PEDIDO", "PEDI_FORMA_PAGO", "PEDI_CODIGO_DSCTO", "PEDI_DESCUENTO_TOTAL", "PEDI_IVA",
                                                    "PEDI_VALOR_IVA", "PEDI_FECHA_POSTERGA_VCTO", "PEDI_FECHA_COBRO", "PEDI_VALOR_DESCUENTO", "PEDI_FECHA_ANULACION", "PEDI_ESTADO",
                                                    "PEDI_USUARIO", "PEDI_TERMINAL", "PEDI_FECHA_SISTEMA", "PEDI_OBSERVACION", "PEDI_ALFA", "PEDI_TOTAL_ICE"
                                                };


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

        //Insert data in the oracle
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
                        string query = $"INSERT INTO {tableName} ({string.Join(",", dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName))}) " +
                                       $"VALUES ({string.Join(",", dataTable.Columns.Cast<DataColumn>().Select(c => $":{c.ColumnName}"))})";

                        using (OracleCommand cmd = new OracleCommand(query, conn))
                        {
                            foreach (DataColumn col in dataTable.Columns)
                            {
                                object value = row[col];

                                // Handle date columns
                                if (col.ColumnName.Equals("PEDI_FECHA", StringComparison.OrdinalIgnoreCase) ||
                                    col.ColumnName.Equals("PEDI_FECHA_ENTREGA", StringComparison.OrdinalIgnoreCase) ||
                                    col.ColumnName.Equals("PEDI_FECHA_SISTEMA", StringComparison.OrdinalIgnoreCase) ||
                                    col.ColumnName.Equals("DEPE_FECHA_ENTREGA", StringComparison.OrdinalIgnoreCase))
                                {
                                    if (value is string stringValue)
                                    {
                                        string[] dateFormats = new string[] { "yyyy-MM-dd", "dd-MM-yyyy", "MM/dd/yyyy", "dd-MM-yy" };
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
                                    else if (value is DateTime dateVal)
                                    {
                                        value = dateVal;  // Keep the DateTime as is if it's already in DateTime format
                                    }
                                }
                                else
                                {
                                    // Handle potential overflow for number or string columns
                                    if (value is decimal decimalValue)
                                    {
                                        // Round decimal to 2 decimal places
                                        decimalValue = Math.Round(decimalValue, 2);

                                        // Adjust precision and scale based on your schema
                                        if (decimalValue > 99999999.99m)  // Example threshold for maximum value
                                        {
                                            WriteToFile($"Value too large for column {col.ColumnName}: {decimalValue}. Truncating value.");
                                            value = 99999999.99m;  // Truncate to max allowed value
                                        }
                                        else
                                        {
                                            value = decimalValue;
                                        }
                                    }
                                    else if (value is string stringVal)
                                    {
                                        // Check for string length (example: 255 characters max for VARCHAR2)
                                        if (stringVal.Length > 255)
                                        {
                                            WriteToFile($"String too long for column {col.ColumnName}: {stringVal}. Truncating value.");
                                            value = stringVal.Substring(0, 255);  // Truncate string to 255 characters
                                        }
                                    }

                                    // Handle null or empty values
                                    if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                                    {
                                        value = DBNull.Value;
                                    }
                                }

                                // Add parameters for the query
                                if (value is DateTime dateValueParam)
                                {
                                    cmd.Parameters.Add(new OracleParameter(col.ColumnName, OracleDbType.Date)).Value = dateValueParam;
                                }
                                else if (value is decimal decimalValueParam)
                                {
                                    cmd.Parameters.Add(new OracleParameter(col.ColumnName, OracleDbType.Decimal)).Value = decimalValueParam;
                                }
                                else
                                {
                                    cmd.Parameters.Add(new OracleParameter(col.ColumnName, value ?? DBNull.Value));
                                }
                            }

                            // Execute the query
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile($"Error inserting data into {tableName}: {ex.Message}");
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
