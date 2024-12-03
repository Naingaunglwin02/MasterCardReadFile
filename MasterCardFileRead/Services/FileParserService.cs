using MasterCardFileRead.Models;
using OfficeOpenXml;
using MasterCardFileRead.Services;

public class FileParserService
{
    public void AddHeaders(ExcelWorksheet worksheet, string[] headers)
    {
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cells[1, i + 1].Value = headers[i];
        }

        // Make the header row bold
        worksheet.Cells[1, 1, 1, headers.Length].Style.Font.Bold = true;

        for (int i = 1; i <= headers.Length; i++)
        {
            worksheet.Column(i).Width = 30;
        }
    }
    public void GenerateExcelFile(List<TransactionModel> headerRecords, List<TransactionModel> feeRecords, string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            var headerSheetRecords = headerRecords
                .Where(record => record.TranscFunction?.Contains("FIRST PRES", StringComparison.OrdinalIgnoreCase) == true)
                .ToList();

            var summarySheetRecords = feeRecords
                .Where(record => record.TranscFunction?.Contains("FEE COLL-CSG", StringComparison.OrdinalIgnoreCase) == true)
                .ToList();

            // Add Header Records sheet
            var ecommerceTransactionSheet = package.Workbook.Worksheets.Add("Acquiring Ecommerce Transaction");
            EcommerceTransaction ecommerceTransaction = new EcommerceTransaction();
            ecommerceTransaction.AddDataToSheet(ecommerceTransactionSheet, headerSheetRecords);

            // Add Summary sheet
            var otherTransactionSheet = package.Workbook.Worksheets.Add("Other Transaction");
            if (summarySheetRecords != null && summarySheetRecords.Any())
            {
                OtherTransaction otherTransaction = new OtherTransaction();
                otherTransaction.AddSummaryDataToSheet(otherTransactionSheet, summarySheetRecords);
            }
            else
            {
                // Optionally add a message or leave it blank if no data is available
                otherTransactionSheet.Cells[1, 1].Value = "No data available for summary.";
            }

            FileInfo fileInfo = new FileInfo(filePath);
            package.SaveAs(fileInfo);
        }
    }
}
