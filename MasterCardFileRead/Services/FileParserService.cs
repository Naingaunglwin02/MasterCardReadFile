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
    }
<<<<<<< HEAD
    public void GenerateExcelFile(List<TransactionModel> ecommerceTransactionRecord, List<TransactionModel> otherTransactionRecord, List<TransactionModel> issuingTransactionRecord, List<RejectTransactionModel> rejectTransactionRecord, string filePath)
=======
    public void GenerateExcelFile(List<TransactionModel> ecommerceTransactionRecord, List<TransactionModel> otherTransactionRecord, List<TransactionModel> posTransactionRecord,  string filePath)
>>>>>>> features/hhz
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage())
        {
            //var filteredEcommerceRecord = ecommerceTransactionRecord
            //    .Where(record => record.MemberID?.Contains("00000017046", StringComparison.OrdinalIgnoreCase) == true).ToList();

            // Filter records by MemberID and FileID's last digit being even
            var filteredEcommerceRecord = ecommerceTransactionRecord
<<<<<<< HEAD
                .Where(record => record.MemberID?.Contains("00000017046", StringComparison.OrdinalIgnoreCase) == true).ToList();

            var filteredOtherRecord = otherTransactionRecord
                .Where(record => record.MemberID?.Contains("00000017046", StringComparison.OrdinalIgnoreCase) == true).ToList();

            var filteredIssuingRecord = issuingTransactionRecord
                .Where(record => record.MemberID?.Contains("00000014688", StringComparison.OrdinalIgnoreCase) == true).ToList();
=======
                .Where(record =>
                    record.MemberID?.Contains("00000017046", StringComparison.OrdinalIgnoreCase) == true &&
                    !string.IsNullOrEmpty(record.FileId) &&
                    FileReadConditionService.ExtractFileIDEven(record.FileId) != null 
                ).ToList();
>>>>>>> features/hhz

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
            otherTransaction.AddSummaryDataToSheet(otherTransactionSheet, filteredOtherRecord);

            var issuingTransactionSheet = package.Workbook.Worksheets.Add("Issuing");
            IssuingTransaction issuingTransaction = new IssuingTransaction();
            issuingTransaction.AddIssuingDataToExcel(issuingTransactionSheet, filteredIssuingRecord);

            var rejectTransactionSheet = package.Workbook.Worksheets.Add("Reject");
            RejectTransaction rejectTransaction = new RejectTransaction();
            rejectTransaction.AddRejectDataToSheet(rejectTransactionSheet, rejectTransactionRecord);

            FileInfo fileInfo = new FileInfo(filePath);
            package.SaveAs(fileInfo);
        }
    }
}
