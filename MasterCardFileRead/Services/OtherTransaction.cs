using MasterCardFileRead.Models;
using OfficeOpenXml;

namespace MasterCardFileRead.Services
{
    public class OtherTransaction
    {
        public List<TransactionModel> ParseFileFee(string filePath)
        {
            var otherTransactionRecords = new List<TransactionModel>();

            using (var reader = new StreamReader(filePath))
            {
                string line;
                string date = null, memberID = null, cycle = null, fileId = null;

                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith("1IP"))
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
                        fileId = FileReadConditionService.ExtractFileIDForOtherTransaction(line);
                    }

                    if (line.Contains("FEE COL CR") || line.Contains("FEE COL DR"))
                    {
                        var otherTransactionResult = FileReadConditionService.ProcessOtherTransaction(line);

                        if (otherTransactionResult != null)
                        {
                            var transaction = new TransactionModel
                            {
                                Date = date,
                                MemberID = memberID,
                                Cycle = cycle,
                                Proc = otherTransactionResult.Proc,
                                FileId = fileId,
                                TranscFunction = otherTransactionResult.TransactionFunction,
                                Code = otherTransactionResult.Code,
                                Count = otherTransactionResult.Count,
                                ReconAmount = otherTransactionResult.ReconAmount,
                                ReconDCDR = otherTransactionResult.ReconDCDR,
                                Currency = otherTransactionResult.Currency,
                                TransferFee = otherTransactionResult.TransferFee,
                                TransferFeeDCDR = otherTransactionResult.TransferFeeDCDR
                            };

                            otherTransactionRecords.Add(transaction);
                        }
                    }

                    //if (line.Contains("***END OF REPORT***"))
                    //{
                    //    break;
                    //}
                }
            }

            return otherTransactionRecords;
        }

        public void AddSummaryDataToSheet(ExcelWorksheet worksheet, List<TransactionModel> otherTransactionRecords)
        {

            string[] headers = new string[]
             {
               "Transaction Function",
                "Date",
                "File ID",
                "Member ID",
                "Cycle",
                "Proc",
                "Code",
                "IRD Values",
                "Count",
                "Recon Amount",
                "DC/CR",
                "Currency",
                "Transfer Fee",
                "DC/CR"
            };


            FileParserService fileParserService = new FileParserService();
            fileParserService.AddHeaders(worksheet, headers);

            int rowIndex = 2;

            // Add data
            foreach (var record in otherTransactionRecords)
            { 
                if (record.EndOfReport == "END")
                {
                    // Insert a blank row
                    rowIndex++;
                    continue;
                }

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
                worksheet.Cells[rowIndex, 11].Value = record.ReconDCDR;
                worksheet.Cells[rowIndex, 12].Value = record.Currency;
                worksheet.Cells[rowIndex, 13].Value = record.TransferFee;
                worksheet.Cells[rowIndex, 14].Value = record.TransferFeeDCDR;

                // Wrap text for multiple-line values
                worksheet.Cells.AutoFitColumns();

                rowIndex++;
            }
        }
    }
}
