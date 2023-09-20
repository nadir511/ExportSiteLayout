using IronXL;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using SixLabors.ImageSharp.PixelFormats;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ExportSiteLayout
{
    public static class ExportSiteLayout
    {
        [FunctionName("ExportSiteLayout")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            try
            {
                log.LogInformation("C# HTTP trigger function processed a request.");

                Console.WriteLine($"File creation starts at: '{DateTime.Now}");
                #region|Common Fields|
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var package = new ExcelPackage();
                #endregion

                #region|Create sheets|
                TenantSheet(package);
                SiteSheet(package);
                AreasSheet(package);
                AssetTypeSheet(package);
                AssetCategorySheet(package);
                AssetSheet(package);
                //SourceTagsSheet(package);
                CreateRealTimeTagSheet(package);
                CreateManualTagSheet(package);
                CreateCalculatedTagSheet(package);
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
            catch (Exception)
            {

                throw;
            }
        }

        static void PopulateWorksheet(ExcelWorksheet worksheet, DataTable dataTable)
        {
            worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
        }
        static void TenantSheet(ExcelPackage package)
        {
            try
            {
                DataTable tenantDataTable = ExecuteStoredProcedure("customerInfoForSiteLayoutBySiteId", 917);
                // Add asset type sheet to the Excel package
                var tenantSheet = package.Workbook.Worksheets.Add("Tenant");

                //Apply header colors
                CreateSheetHeader(tenantSheet);

                // Populate the Asset Type data
                PopulateWorksheet(tenantSheet, tenantDataTable);
            }
            catch (Exception)
            {

                throw;
            }
        }
        static void SiteSheet(ExcelPackage package)
        {
            try
            {
                DataTable siteDataTable = ExecuteStoredProcedure("siteInfoForSiteLayoutBySiteId", 917);

                // Add asset type sheet to the Excel package
                var siteSheet = package.Workbook.Worksheets.Add("Site");

                //Apply header colors
                CreateSheetHeader(siteSheet);

                // Populate the Asset Type data
                PopulateWorksheet(siteSheet, siteDataTable);
            }
            catch (Exception)
            {

                throw;
            }
        }
        static void AreasSheet(ExcelPackage package)
        {
            try
            {
                DataTable areaDataTable = ExecuteStoredProcedure("areaInfoForSiteLayoutBySiteId", 917);

                // Add asset type sheet to the Excel package
                var areaSheet = package.Workbook.Worksheets.Add("Areas");

                //Apply header colors
                CreateSheetHeader(areaSheet);

                // Create a data validation list for the third column (column C/Action Type) starting from the second row (row 2)
                AddActionTypeColumn(areaSheet, areaDataTable);

                //Add required filed color as red
                areaSheet.Cells["B1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                areaSheet.Cells["C1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                // Populate the Asset Type data
                PopulateWorksheet(areaSheet, areaDataTable);
            }
            catch (Exception)
            {

                throw;
            }
        }
        static void AssetTypeSheet(ExcelPackage package)
        {
            try
            {
                var sqlQuery = "SELECT  AST.AT_ID,\r\nAST.AT_Name AS AssetTypeName,\r\nAST.AT_Description AS Description\r\nFROM ASSET_TYPE AST";
                DataTable assetTypeDataTable = ExecuteSqlQuery(sqlQuery);

                // Add asset type sheet to the Excel package
                var assetTypeSheet = package.Workbook.Worksheets.Add("Asset Types");

                //Apply header colors
                CreateSheetHeader(assetTypeSheet);
                
                // Populate the Asset Type data
                PopulateWorksheet(assetTypeSheet, assetTypeDataTable);
            }
            catch (Exception)
            {

                throw;
            }
        }
        static void AssetCategorySheet(ExcelPackage package)
        {
            try
            {
                var sqlQuery = "SELECT\r\nAC.AC_Name,\r\nAST.AT_Name,\r\nAC.AC_Description\r\nFROM ASSET_TYPE AST INNER JOIN ASSET_CATEGORY AC ON AST.AT_ID=AC.AT_ID";
                DataTable assetTypeDataTable = ExecuteSqlQuery(sqlQuery);

                // Add asset type sheet to the Excel package
                var assetTypeSheet = package.Workbook.Worksheets.Add("Asset Categories");
                //Apply header colors
                CreateSheetHeader(assetTypeSheet);
                // Populate the Asset Type data
                PopulateWorksheet(assetTypeSheet, assetTypeDataTable);
            }
            catch (Exception)
            {

                throw;
            }
        }
        static void AssetSheet(ExcelPackage package)
        {
            try
            {
                DataTable assetDataTable = ExecuteStoredProcedure("assetListForSiteLayoutBySiteId", 917);

                // Add asset type sheet to the Excel package
                var assetTypeSheet = package.Workbook.Worksheets.Add("Assets");
                //Apply header colors
                CreateSheetHeader(assetTypeSheet);

                // Create a data validation list for the third column (column C/Action Type) starting from the second row (row 2)
                AddActionTypeColumn(assetTypeSheet, assetDataTable);

                //Add required filed color as red
                assetTypeSheet.Cells["B1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                assetTypeSheet.Cells["C1"].Value = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                assetTypeSheet.Cells["D1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                assetTypeSheet.Cells["E1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                assetTypeSheet.Cells["F1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                assetTypeSheet.Cells["G1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                //Add comments on header fields
                ExcelRange Rng0 = assetTypeSheet.Cells["E1"];
                Rng0.Value = "Asset Classification";
                Rng0.AddComment("It can be 'Group' or 'Asset' ", "OmniConnect");

                // Populate the Asset Type data
                PopulateWorksheet(assetTypeSheet, assetDataTable);
            }
            catch (Exception)
            {

                throw;
            }
        }
        static void SourceTagsSheet(ExcelPackage package)
        {
            try
            {
                DataTable sourceTagDataTable = ExecuteStoredProcedure("getSourceTagAndDeviceInfoBySiteId", 917);

                // Add source tag sheet to the Excel package
                var sourceTagSheet = package.Workbook.Worksheets.Add("Source Tags");

                //Apply header colors
                CreateSheetHeader(sourceTagSheet);

                //Add required filed color as red

            }
            catch
            {

            }
        }
        static void CreateRealTimeTagSheet(ExcelPackage package)
        {
            try
            {
                DataTable realTagDataTable = ExecuteStoredProcedure("rawTagsListForSiteLayoutBySiteId", 917);

                // Add real tag sheet to the Excel package
                var realTimeTagSheet = package.Workbook.Worksheets.Add("Real Time Tags");

                
                //Apply header colors
                CreateSheetHeader(realTimeTagSheet);

                // Create a data validation list for the third column (column C/Action Type) starting from the second row (row 2)
                AddActionTypeColumn(realTimeTagSheet, realTagDataTable);

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

                
                ExcelRange Rng2 = realTimeTagSheet.Cells["F1"];
                Rng2.Value = "Device Type";
                Rng2.AddComment("The Device Type can be 'OPC Device' or 'IOT Device' or 'Modbus Device' ", "OmniConnect");

                // Populate the Real time data
                PopulateWorksheet(realTimeTagSheet, realTagDataTable);
            }
            catch (Exception)
            {

                throw;
            }
        }
        static void CreateManualTagSheet(ExcelPackage package)
        {
            try
            {
                DataTable manualTagDataTable = ExecuteStoredProcedure("manualTagsListForSiteLayoutBySiteId", 917, "e5fc6d55-c68c-4471-8aef-e43cc011c233");

                // Add real tag sheet to the Excel package
                var manualTagSheet = package.Workbook.Worksheets.Add("Manual Tags");


                //Apply header colors
                CreateSheetHeader(manualTagSheet);

                // Create a data validation list for the third column (column C/Action Type) starting from the second row (row 2)
                AddActionTypeColumn(manualTagSheet, manualTagDataTable);

                manualTagSheet.Cells["B1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                manualTagSheet.Cells["D1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                // Populate the manual tag data
                PopulateWorksheet(manualTagSheet, manualTagDataTable);

            }
            catch (Exception)
            {

                throw;
            }
        }
        static void CreateCalculatedTagSheet(ExcelPackage package)
        {
            try
            {
                DataTable calTagDataTable = ExecuteStoredProcedure("calTagsListForSiteLayoutBySiteId", 917, "e5fc6d55-c68c-4471-8aef-e43cc011c233");

                // Add real tag sheet to the Excel package
                var calTagSheet = package.Workbook.Worksheets.Add("Calculated Tags");


                //Apply header colors
                CreateSheetHeader(calTagSheet);

                // Create a data validation list for the third column (column C/Action Type) starting from the second row (row 2)
                AddActionTypeColumn(calTagSheet, calTagDataTable);

                calTagSheet.Cells["B1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                calTagSheet.Cells["D1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                calTagSheet.Cells["H1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                // Populate the manual tag data
                PopulateWorksheet(calTagSheet, calTagDataTable);
            }
            catch (Exception)
            {

                throw;
            }
        }
        static void CreateSheetHeader(ExcelWorksheet sheet)
        {
            try
            {
                // Hide the ID Column
                sheet.Column(1).Hidden = true;

                // Set the first row as a header
                sheet.Row(1).Style.Font.Bold = true;
                sheet.Row(1).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Row(1).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                // Freeze the first row
                sheet.View.FreezePanes(2, 1);
            }
            catch (Exception)
            {

                throw;
            }
        }
        static void AddActionTypeColumn(ExcelWorksheet sheet,DataTable dataTable)
        {
            try
            {
                sheet.Cells["C1"].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                var dataValidation = sheet.Cells["C2:C" + (dataTable.Rows.Count + 1)].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                dataValidation.Formula.Values.Add("Update");
                dataValidation.Formula.Values.Add("Add");

                ExcelRange Rng1 = sheet.Cells["C1"];
                Rng1.Value = "Action Type";
                Rng1.AddComment("It can be 'Update' or 'Add' ", "OmniConnect");
            }
            catch (Exception)
            {

                throw;
            }
        }
        static DataTable ExecuteStoredProcedure(string procedureName, int siteId,string userId=null)
        {
            try
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
                        if (userId!=null)
                        {
                            command.Parameters.AddWithValue("@userId", userId);
                        }

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);
                            return dataTable;
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        static DataTable ExecuteSqlQuery(string sqlQuery)
        {
            try
            {
                string connectionString = Environment.GetEnvironmentVariable("ConnectionString");
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);
                            return dataTable;
                        }
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}

