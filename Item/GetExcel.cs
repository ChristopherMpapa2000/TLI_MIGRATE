using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigrateTLI.Item
{
    class GetExcel
    {
        public static Dictionary<string, DataTable> ReadExcelToDataTables(string filePath)
        {
            var result = new Dictionary<string, DataTable>();

            // สร้าง Connection String สำหรับ Excel
            string connectionString = GetExcelConnectionString(filePath);

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    // เปิดการเชื่อมต่อ
                    connection.Open();

                    // ดึง Schema เพื่อดูรายชื่อ Sheets
                    DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    // อ่านข้อมูลจากแต่ละ Sheet
                    foreach (DataRow sheet in schemaTable.Rows)
                    {
                        string sheetName = sheet["TABLE_NAME"].ToString();

                        // อ่านเฉพาะ Sheet ที่ลงท้ายด้วย '$' (เป็นชื่อ Sheet จริง)
                        if (sheetName.Trim('\'').EndsWith("$"))
                        {
                            string query = $"SELECT * FROM [{sheetName}]";

                            using (var adapter = new OleDbDataAdapter(query, connection))
                            {
                                var dataTable = new DataTable(sheetName.Trim('\'').TrimEnd('$'));
                                adapter.Fill(dataTable);

                                foreach (DataRow row in dataTable.Rows)
                                {
                                    for (int col = 0; col < dataTable.Columns.Count; col++)
                                    {
                                        if (row[col] == DBNull.Value || row[col] == null || row[col] == "NULL")
                                        {
                                            row[col] = ""; // แทนที่ด้วยค่าว่าง
                                        }
                                    }
                                }

                                // เพิ่ม DataTable ลงใน Dictionary
                                result[sheetName.Trim('\'').TrimEnd('$')] = dataTable;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                    Program.LogError("ReadExcelToDataTables: " + ex);
                    Environment.Exit(1);
                }
            }

            return result;
        }
        private static string GetExcelConnectionString(string filePath)
        {
            string extension = System.IO.Path.GetExtension(filePath);

            if (extension == ".xls")
            {
                return $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            }
            else if (extension == ".xlsx")
            {
                return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1\"";
            }
            else
            {
                throw new NotSupportedException("Unsupported file format");
            }
        }
    }
}
