using MasterCardFileRead.Models;
using OfficeOpenXml;
using MasterCardFileRead.Services;
using System.Text;
using MasterCardFileRead.Constants;

namespace MasterCardFileRead.Services
{
    public class PosTransaction
    {
        public List<TransactionModel> ParseFilePos(string filePath)
        {
            var posTransactionRecords = new List<TransactionModel>();

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (var reader = new StreamReader(filePath, Encoding.GetEncoding(CommonConstants.Encoding)))
            {
                string line;
                string date = null, memberID = null, cycle = null, fileId = null, endOfReport = null;

                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Contains(CommonConstants.Date))
                    {
                        date = FileReadConditionService.ExtractDate(line, ref date);
                    }

                    if (line.Contains(CommonConstants.MemberId))
                    {
                        memberID = FileReadConditionService.ExtractMemberID(line);
                    }

                    if (line.Contains(CommonConstants.Cycle))
                    {
                        cycle = FileReadConditionService.ExtractAcceptanceBrandCycle(line);
                    }

                    if (line.Contains(CommonConstants.FileId))
                    {
                        fileId = FileReadConditionService.ExtractFileID(line);
                    }

                    var posTransactionResult = FileReadConditionService.ProcessIssuingTransaction(line);

                    if (posTransactionResult != null)
                    {
                        var transaction = new TransactionModel
                        {
                            Date = date,
                            MemberID = memberID,
                            Cycle = cycle,
                            Proc = posTransactionResult.Proc,
                            FileId = fileId,
                            TranscFunction = posTransactionResult.TransactionFunction,
                            Code = posTransactionResult.Code,
                            Ird = posTransactionResult.IrdValues,
                            Count = posTransactionResult.Count,
                            ReconAmount = posTransactionResult.ReconAmount,
                            ReconDCCR = posTransactionResult.ReconDCCR,
                            Currency = posTransactionResult.Currency,
                            TransferFee = posTransactionResult.TransferFee,
                            TransferFeeDCCR = posTransactionResult.TransferFeeDCCR,
                        };

                        posTransactionRecords.Add(transaction);
                    }
                }
            }

            return posTransactionRecords;
        }

        public void AddDataToSheet(ExcelWorksheet worksheet, List<TransactionModel> ecommerceTransactionRecords)
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
            string previousFileId = null;


            int totalCount = 0;
            double totalDr = 0, totalCr = 0, totalTranDr = 0, totalTranCr = 0, totalTransFee = 0;
            int grandTotalCount = 0;
            double grandTotalDr = 0, grandTotalCr = 0, grandTotalTransDr = 0, grandTotalTransCr = 0;

            foreach (var record in ecommerceTransactionRecords)
            {
                if (CalculateDrCr.ShouldAddSubtotals(previousFileId, record.FileId))
                {
                    CalculateDrCr.AddSubtotals(worksheet, ref rowIndex, ref totalCount, ref totalDr, ref totalCr, ref totalTranDr, ref totalTranCr, ref grandTotalDr, ref grandTotalCr, ref grandTotalTransDr, ref grandTotalTransCr);
                }

                if (CalculateDrCr.ShouldAddGrandTotals(previousDate, record.Date))
                {
                    // CalculateDrCr.AddSubtotals(worksheet, ref rowIndex, ref totalCount, ref totalDr, ref totalCr, ref totalTranDr, ref totalTranCr, ref grandTotalDr, ref grandTotalCr, ref grandTotalTransDr, ref grandTotalTransCr);
                    CalculateDrCr.AddGrandTotals(worksheet, ref rowIndex, grandTotalCount, grandTotalDr, grandTotalCr, grandTotalTransDr, grandTotalTransCr);
                    TotalTransactions.ResetGrandTotalVariables(ref grandTotalCount, ref grandTotalDr, ref grandTotalCr, ref grandTotalTransDr, ref grandTotalTransCr);
                }

                CalculateDrCr.WriteTransactionData(worksheet, record, rowIndex, ref totalDr, ref totalCr, ref totalTranDr, ref totalTranCr, ref totalCount, ref grandTotalCount);

                previousCycle = record.Cycle;
                previousDate = record.Date;
                previousFileId = record.FileId;
                rowIndex++;
            }

            CalculateDrCr.AddFinalSubtotalsAndGrandTotals(worksheet, ref rowIndex, previousCycle, totalCount, totalDr, totalCr, totalTranDr, totalTranCr, ref grandTotalDr, ref grandTotalCr, ref grandTotalTransDr, ref grandTotalTransCr, grandTotalCount);
        }

    }
}
