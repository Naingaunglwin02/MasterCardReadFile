using MasterCardFileRead.Models;
using OfficeOpenXml;
using MasterCardFileRead.Services;

namespace MasterCardFileRead.Services
{
    public class PosTransaction
    {
        public List<TransactionModel> ParseFilePos(string filePath)
        {
            var posTransactionRecords = new List<TransactionModel>();

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

                    if (line.Contains("PURCHASE"))
                    {
                        var posTransactionResult = FileReadConditionService.ProcessEcommerceTransaction(line);

                        if (posTransactionResult != null)
                        {
                            var transaction = new TransactionModel
                            {
                                Date = date,
                                MemberID = memberID,
                                Cycle = cycle,
                                FileId = fileId,
                                TranscFunction = posTransactionResult.TransactionFunction,
                                Code = posTransactionResult.Code,
                                Ird = posTransactionResult.IrdValues,
                                Count = posTransactionResult.Count,
                                ReconAmount = posTransactionResult.ReconAmount,
                                TransferFee = posTransactionResult.TransferFee
                            };

                            posTransactionRecords.Add(transaction);
                        }
                    }

                    //if (line.Contains("***END OF REPORT***"))
                    //{
                    //    break;
                    //}
                }
            }

            return posTransactionRecords;
        }

        public void AddDataToSheet(ExcelWorksheet worksheet, List<TransactionModel> posTransactionRecords)
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
            foreach (var record in posTransactionRecords)
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
