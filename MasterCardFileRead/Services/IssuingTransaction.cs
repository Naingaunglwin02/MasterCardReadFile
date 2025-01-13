using MasterCardFileRead.Models;
using OfficeOpenXml;

namespace MasterCardFileRead.Services
{
    public class IssuingTransaction
    {
        public List<TransactionModel> IssuingTransactionService(string filePath)
        {
            var issuingTransactionRecords = new List<TransactionModel>();

            using (var reader = new StreamReader(filePath))
            {
                string line;
                string date = null, memberID = null, cycle = null, fileId = null, endOfReport = null;

                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Contains("BUSINESS SERVICE LEVEL:"))
                    {
                        date = FileReadConditionService.ExtractDate(line, ref date);
                    }

                    if (line.Contains("MEMBER ID:"))
                    {
                        memberID = FileReadConditionService.ExtractMemberID(line);
                    }

                    if (line.Contains("ACCEPTANCE BRAND:"))
                    {
                        cycle = FileReadConditionService.ExtractAcceptanceBrandCycle(line);
                    }

                    if (line.Contains("FILE ID:"))
                    {
                        fileId = FileReadConditionService.ExtractFileID(line);
                    }

                    var ecommerceTransactionResult = FileReadConditionService.ProcessIssuingTransaction(line);

                    if (ecommerceTransactionResult != null)
                    {
                        var transaction = new TransactionModel
                        {
                            Date = date,
                            MemberID = memberID,
                            Cycle = cycle,
                            Proc = ecommerceTransactionResult.Proc,
                            FileId = fileId,
                            TranscFunction = ecommerceTransactionResult.TransactionFunction,
                            Code = ecommerceTransactionResult.Code,
                            Ird = ecommerceTransactionResult.IrdValues,
                            Count = ecommerceTransactionResult.Count,
                            ReconAmount = ecommerceTransactionResult.ReconAmount,
                            ReconDCCR = ecommerceTransactionResult.ReconDCCR,
                            Currency = ecommerceTransactionResult.Currency,
                            TransferFee = ecommerceTransactionResult.TransferFee,
                            TransferFeeDCCR = ecommerceTransactionResult.TransferFeeDCCR,
                        };

                        issuingTransactionRecords.Add(transaction);
                    }
                }
            }

            return issuingTransactionRecords;
        }

        public void AddIssuingDataToExcel(ExcelWorksheet worksheet, List<TransactionModel> ecommerceTransactionRecords)
        {
            string[] headers = new string[]
            {
                "TRANS FUNC",
                "DATE",
                "FILE ID",
                "MEMBER ID",
                "CYCLE",
                "PROC",
                "CODE",
                "IRD",
                "COUNT",
                "RECON AMOUNT",
                "DR/ CR",
                "CURRENCY",
                "TRANS FEE",
                "DR/ CR"
            };

            FileParserService fileParserService = new FileParserService();
            fileParserService.AddHeaders(worksheet, headers, 15);

            int rowIndex = 2;

            string previousCycle = null;
            string previousDate = null;

            // Subtotal variables
            int totalCount = 0;
            double totalRecon = 0;
            double totalTransFee = 0;
            string totalCr = "";
            string totalDr = "";

            // Grand total variables
            int grandTotalCount = 0;
            double grandTotalRecon = 0;
            double grandTotalTransFee = 0;

            // Add data
            foreach (var record in ecommerceTransactionRecords)
            {
                // Add subtotals if the cycle changes
                if (previousCycle != null && record.Cycle != previousCycle)
                {
                    TotalTransactions.AddSubtotalRow(worksheet, rowIndex, "Total", totalCount, totalRecon, totalTransFee, totalCr, totalDr);
                    rowIndex += 2;
                    TotalTransactions.ResetSubtotalVariables(ref totalCount, ref totalRecon, ref totalTransFee, ref totalCr, ref totalDr);
                }

                // Add grand totals if the date changes
                if (previousDate != null && record.Date != previousDate)
                {
                    TotalTransactions.AddGrandTotalRow(worksheet, rowIndex, "Grand Total", grandTotalCount, grandTotalRecon, grandTotalTransFee);
                    rowIndex += 2;
                    TotalTransactions.ResetGrandTotalVariables(ref grandTotalCount, ref grandTotalRecon, ref grandTotalTransFee);
                }

                // Write transaction data
                worksheet.Cells[rowIndex, 1].Value = record.TranscFunction;
                worksheet.Cells[rowIndex, 2].Value = record.Date;
                worksheet.Cells[rowIndex, 3].Value = record.FileId;
                worksheet.Cells[rowIndex, 4].Value = record.MemberID;
                worksheet.Cells[rowIndex, 5].Value = record.Cycle;
                worksheet.Cells[rowIndex, 6].Value = record.Proc;
                worksheet.Cells[rowIndex, 7].Value = record.Code;
                worksheet.Cells[rowIndex, 8].Value = record.Ird;
                worksheet.Cells[rowIndex, 9].Value = record.Count;
                worksheet.Cells[rowIndex, 10].Value = record.ReconAmount;
                worksheet.Cells[rowIndex, 11].Value = record.ReconDCCR;
                worksheet.Cells[rowIndex, 12].Value = record.Currency;
                worksheet.Cells[rowIndex, 13].Value = record.TransferFee;
                worksheet.Cells[rowIndex, 14].Value = record.TransferFeeDCCR;

                // Update subtotal and grand total variables
                totalCount += Int32.Parse(record.Count);
                totalRecon += Convert.ToDouble(record.ReconAmount);
                totalTransFee += Convert.ToDouble(record.TransferFee);
                totalCr = record.ReconDCCR;
                totalDr = record.TransferFeeDCCR;

                grandTotalCount += Int32.Parse(record.Count);
                grandTotalRecon += Convert.ToDouble(record.ReconAmount);
                grandTotalTransFee += Convert.ToDouble(record.TransferFee);

                previousCycle = record.Cycle;
                previousDate = record.Date;
                rowIndex++;
            }

            // Add final subtotals and grand totals
            if (previousCycle != null)
            {
                TotalTransactions.AddSubtotalRow(worksheet, rowIndex, "Total", totalCount, totalRecon, totalTransFee, totalCr, totalDr);
                rowIndex += 2;
            }

            if (previousDate != null)
            {
                TotalTransactions.AddGrandTotalRow(worksheet, rowIndex, "Grand Total", grandTotalCount, grandTotalRecon, grandTotalTransFee);
                rowIndex += 2;
            }
            worksheet.Cells.AutoFitColumns();
        }
    }
}
