using IronXL;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;

namespace ExportSiteLayout
{
    public static class ExportSiteLayout
    {
        [FunctionName("ExportSiteLayout")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            Console.WriteLine($"File creation starts at: '{DateTime.Now}");
            #region|Common Fields|
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage();
            #endregion

            #region|Create file using EPplus using datatable|

            DataTable realTagDataTable = ExecuteStoredProcedure("rawTagsListForSiteLayoutBySiteId", 917);

            // Add real tag sheet to the Excel package
            var realTimeTagSheet = package.Workbook.Worksheets.Add("Sheet1");

            // Hide the ID Column
            realTimeTagSheet.Column(1).Hidden = true;

            // Set the first row as a header
            realTimeTagSheet.Row(1).Style.Font.Bold = true;
            realTimeTagSheet.Row(1).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            realTimeTagSheet.Row(1).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

            // Freeze the first row
            realTimeTagSheet.View.FreezePanes(2, 1);

            // Create a data validation list for the third column (column C/Action Type) starting from the second row (row 2)
            var dataValidation = realTimeTagSheet.Cells["C2:C" + (realTagDataTable.Rows.Count + 1)].DataValidation.AddListDataValidation() as ExcelDataValidationList;
            dataValidation.Formula.Values.Add("Update");
            dataValidation.Formula.Values.Add("Add");

            //Add required filed color as red
            realTimeTagSheet.Cells["B1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
            realTimeTagSheet.Cells["C1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
            realTimeTagSheet.Cells["D1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
            realTimeTagSheet.Cells["E1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
            realTimeTagSheet.Cells["F1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
            realTimeTagSheet.Cells["K1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);

            //Add comments on header fields
            ExcelRange Rng0 = realTimeTagSheet.Cells["K1"];
            Rng0.Value = "Data Type";
            Rng0.AddComment("The data type can only be 'Boolean', 'Int16', 'UInt16', 'Int32', 'UInt32', 'Int64', 'UInt64', 'Float', 'Double', 'Digital', 'Integer', 'Decimal', 'String'", "OmniConnect");

            ExcelRange Rng1 = realTimeTagSheet.Cells["C1"];
            Rng1.Value = "Action Type";
            Rng1.AddComment("It can be 'Update' or 'Add' ", "OmniConnect");

            ExcelRange Rng2 = realTimeTagSheet.Cells["F1"];
            Rng2.Value = "Device Type";
            Rng2.AddComment("The Device Type can be 'OPC Device' or 'IOT Device' or 'Modbus Device' ", "OmniConnect");

            // Populate the Real time data
            PopulateWorksheet(realTimeTagSheet, realTagDataTable);
            
            #endregion

            #region|Saving the excel file|
            // Save the Excel package to a MemoryStream
            MemoryStream stream = new MemoryStream();
            package.SaveAs(stream);

            // Set the position of the stream back to the beginning
            stream.Seek(0, SeekOrigin.Begin);
            Console.WriteLine($"File creation completed at: '{DateTime.Now}");
            // Return the Excel file as an HTTP response
            return new FileStreamResult(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "MultiSheetExcel.xlsx"
            };
            #endregion
        }

        static void PopulateWorksheet(ExcelWorksheet worksheet, DataTable dataTable)
        {
            worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
        }
        static DataTable ExecuteStoredProcedure(string procedureName, int siteId)
        {
            string connectionString = Environment.GetEnvironmentVariable("ConnectionString");
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(procedureName, connection))
                {
                    command.CommandType = CommandType.StoredProcedure;

                    // Add parameters if necessary
                    command.Parameters.AddWithValue("@SiteId", siteId);

                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);
                        return dataTable;
                    }
                }
            }
        }
    }
}

public class ExportedData
{
    public int Real_Tag_Id { get; set; }
    public string R_Tag_Name { get; set; }
    public string Source_Tag_Name { get; set; }
    public string Name { get; set; }
    public string DeviceTypeName { get; set; }
    public string RT_Unit { get; set; }
    public string DataType { get; set; }
    public double? Publish_Rate { get; set; } // Use appropriate data type for Publish_Rate
    public double? Heart_Beat_Rate { get; set; } // Use appropriate data type for Heart_Beat_Rate
    public double? Scan_Rate { get; set; } // Use appropriate data type for Scan_Rate
    public string Description { get; set; }
}

