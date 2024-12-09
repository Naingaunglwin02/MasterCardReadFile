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
                        fileId = FileReadConditionService.ExtractFileIDEven(line);
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
                                FileId = fileId,
                                TranscFunction = otherTransactionResult.TransactionFunction,
                                Code = otherTransactionResult.Code,
                                Count = otherTransactionResult.Count,
                                ReconAmount = otherTransactionResult.ReconAmount,
                                TransferFee = otherTransactionResult.TransferFee
                            };

                            otherTransactionRecords.Add(transaction);
                        }
                    }

                    if (line.Contains("***END OF REPORT***"))
                    {
                        break;
                    }
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
                "Code",
                "IRD Values",
                "Count",
                "Recon Amount",
                "Transfer Fee"
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
                worksheet.Cells[rowIndex, 6].Value = record.Code;
                worksheet.Cells[rowIndex, 7].Value = record.Ird;
                worksheet.Cells[rowIndex, 8].Value = record.Count;
                worksheet.Cells[rowIndex, 9].Value = record.ReconAmount;
                worksheet.Cells[rowIndex, 10].Value = record.TransferFee;

                // Wrap text for multiple-line values
                worksheet.Cells[rowIndex, 6, rowIndex, 10].Style.WrapText = true;

                rowIndex++;
            }
        }
    }
}
