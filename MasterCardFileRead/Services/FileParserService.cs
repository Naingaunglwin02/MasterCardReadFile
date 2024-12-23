using MasterCardFileRead.Models;
using OfficeOpenXml;
using MasterCardFileRead.Services;

public class FileParserService
{
    public void AddHeaders(ExcelWorksheet worksheet, string[] headers, float fontSize)
    {
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cells[1, i + 1].Value = headers[i];
        }

        // Make the header row bold
        var headerRange = worksheet.Cells[1, 1, 1, headers.Length];
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Font.Size = fontSize;
    }

    public void GenerateExcelFile(List<TransactionModel> ecommerceTransactionRecord, List<TransactionModel> otherTransactionRecord, List<TransactionModel> issuingTransactionRecord, List<TransactionModel> posTransactionRecord, string filePath)

    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage())
        {
            // Filter records by MemberID and FileID's last digit being even
            var filteredEcommerceRecord = ecommerceTransactionRecord
                .Where(record => record.MemberID?.Contains("00000017046", StringComparison.OrdinalIgnoreCase) == true &&
                  !string.IsNullOrEmpty(record.FileId) &&
                  FileReadConditionService.ExtractFileIDEven(record.FileId) != null).ToList();

            var filteredOtherRecord = otherTransactionRecord
                .Where(record => record.MemberID?.Contains("00000017046", StringComparison.OrdinalIgnoreCase) == true &&
                  !string.IsNullOrEmpty(record.FileId) &&
                  FileReadConditionService.ExtractFileIDForOtherTransaction(record.FileId) != null).ToList();

            var filteredIssuingRecord = issuingTransactionRecord
                .Where(record => record.MemberID?.Contains("00000014688", StringComparison.OrdinalIgnoreCase) == true).ToList();

            var filteredPosRecord = posTransactionRecord
                .Where(record =>
                  record.MemberID?.Contains("00000017046", StringComparison.OrdinalIgnoreCase) == true &&
                  !string.IsNullOrEmpty(record.FileId) &&
                  FileReadConditionService.ExtractFileIDOdd(record.FileId) != null).ToList();


            var ecommerceTransactionSheet = package.Workbook.Worksheets.Add("Acquiring_Ecommerce");
            EcommerceTransaction ecommerceTransaction = new EcommerceTransaction();
            ecommerceTransaction.AddDataToSheet(ecommerceTransactionSheet, filteredEcommerceRecord);


            var posTransactionSheet = package.Workbook.Worksheets.Add("Acquiring_Transaction");
            PosTransaction posTransaction = new PosTransaction();
            posTransaction.AddDataToSheet(posTransactionSheet, filteredPosRecord);

            // Add Summary sheet
            var otherTransactionSheet = package.Workbook.Worksheets.Add("Acquiring_Others");
            OtherTransaction otherTransaction = new OtherTransaction();
            otherTransaction.AddSummaryDataToSheet(otherTransactionSheet, filteredOtherRecord);

            var issuingTransactionSheet = package.Workbook.Worksheets.Add("Issuing");
            IssuingTransaction issuingTransaction = new IssuingTransaction();
            issuingTransaction.AddIssuingDataToExcel(issuingTransactionSheet, filteredIssuingRecord);

            //var rejectTransactionSheet = package.Workbook.Worksheets.Add("Reject");
            //RejectTransaction rejectTransaction = new RejectTransaction();
            //rejectTransaction.AddRejectDataToSheet(rejectTransactionSheet, rejectTransactionRecord);

            FileInfo fileInfo = new FileInfo(filePath);
            package.SaveAs(fileInfo);
        }
    }
}
