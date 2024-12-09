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
    public void GenerateExcelFile(List<TransactionModel> ecommerceTransactionRecord, List<TransactionModel> otherTransactionRecord, List<TransactionModel> posTransactionRecord,  string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        using (var package = new ExcelPackage())
        {
            //var filteredEcommerceRecord = ecommerceTransactionRecord
            //    .Where(record => record.MemberID?.Contains("00000017046", StringComparison.OrdinalIgnoreCase) == true).ToList();

            // Filter records by MemberID and FileID's last digit being even
            var filteredEcommerceRecord = ecommerceTransactionRecord
                .Where(record =>
                    record.MemberID?.Contains("00000017046", StringComparison.OrdinalIgnoreCase) == true &&
                    !string.IsNullOrEmpty(record.FileId) &&
                    FileReadConditionService.ExtractFileIDEven(record.FileId) != null 
                ).ToList();

            var ecommerceTransactionSheet = package.Workbook.Worksheets.Add("Ecommerce Transaction");
            EcommerceTransaction ecommerceTransaction = new EcommerceTransaction();
            ecommerceTransaction.AddDataToSheet(ecommerceTransactionSheet, filteredEcommerceRecord);

            //update for pos sheet

            var filteredPosRecord = posTransactionRecord
               .Where(record =>
                   record.MemberID?.Contains("00000017046", StringComparison.OrdinalIgnoreCase) == true &&
                   !string.IsNullOrEmpty(record.FileId) &&
                   FileReadConditionService.ExtractFileIDEven(record.FileId) != null
               ).ToList();

            var posTransactionSheet = package.Workbook.Worksheets.Add("Pos Transaction");
            PosTransaction posTransaction = new PosTransaction();
            posTransaction.AddDataToSheet(posTransactionSheet, filteredPosRecord);

            // Add Summary sheet
            var otherTransactionSheet = package.Workbook.Worksheets.Add("Other Transaction");
            OtherTransaction otherTransaction = new OtherTransaction();
            otherTransaction.AddSummaryDataToSheet(otherTransactionSheet, otherTransactionRecord);
          
            FileInfo fileInfo = new FileInfo(filePath);
            package.SaveAs(fileInfo);
        }
    }
}
