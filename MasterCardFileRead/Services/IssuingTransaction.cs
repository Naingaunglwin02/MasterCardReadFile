using MasterCardFileRead.Models;
using OfficeOpenXml;
using System.IO;
using System.Text;

namespace MasterCardFileRead.Services
{
    public class IssuingTransaction
    {
        public List<TransactionModel> IssuingTransactionService(string filePath)
        {
            var issuingTransactionRecords = new List<TransactionModel>();

            using (var reader = new StreamReader(filePath, Encoding.GetEncoding("Windows-1252")))
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

            int totalCount = 0;
            double totalDr = 0, totalCr = 0, totalTranDr = 0, totalTranCr = 0, totalTransFee = 0;
            int grandTotalCount = 0;
            double grandTotalDr = 0, grandTotalCr = 0, grandTotalTransDr = 0, grandTotalTransCr = 0;

            foreach (var record in ecommerceTransactionRecords)
            {
                if (CalculateDrCr.ShouldAddSubtotals(previousCycle, record.Cycle))
                {
                    CalculateDrCr.AddSubtotals(worksheet, ref rowIndex, ref totalCount, ref totalDr, ref totalCr, ref totalTranDr, ref totalTranCr, ref grandTotalDr, ref grandTotalCr, ref grandTotalTransDr, ref grandTotalTransCr);
                }

                if (CalculateDrCr.ShouldAddGrandTotals(previousDate, record.Date))
                {
                    CalculateDrCr.AddGrandTotals(worksheet, ref rowIndex, grandTotalCount, grandTotalDr, grandTotalCr, grandTotalTransDr, grandTotalTransCr);
                    TotalTransactions.ResetGrandTotalVariables(ref grandTotalCount, ref grandTotalDr, ref grandTotalCr, ref grandTotalTransDr, ref grandTotalTransCr);
                }

                CalculateDrCr.WriteTransactionData(worksheet, record, rowIndex, ref totalDr, ref totalCr, ref totalTranDr, ref totalTranCr, ref totalCount, ref grandTotalCount);

                previousCycle = record.Cycle;
                previousDate = record.Date;
                rowIndex++;
            }

            CalculateDrCr.AddFinalSubtotalsAndGrandTotals(worksheet, ref rowIndex, previousCycle, totalCount, totalDr, totalCr, totalTranDr, totalTranCr, ref grandTotalDr, ref grandTotalCr, ref grandTotalTransDr, ref grandTotalTransCr, grandTotalCount);
        }
    }
}
