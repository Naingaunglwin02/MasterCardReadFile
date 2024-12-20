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

                    if (line.Contains("***END OF REPORT***"))
                    {
                        endOfReport = FileReadConditionService.ExtractEndOfReport(line);
                        System.Diagnostics.Debug.WriteLine(endOfReport, "this is end of report.....");
                    }

                    //if (line.Contains("PURCHASE") || line.Contains("FEE COL CR") || line.Contains("FEE COL DR"))
                    //{
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
                            //EndOfReport = endOfReport
                        };

                        issuingTransactionRecords.Add(transaction);
                        //}
                    }

                    //if (line.Contains("***END OF REPORT***"))
                    //{
                    //    break;
                    //}
                }
            }

            return issuingTransactionRecords;
        }

        public void AddIssuingDataToExcel(ExcelWorksheet worksheet, List<TransactionModel> ecommerceTransactionRecords)
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
            fileParserService.AddHeaders(worksheet, headers, 15);

            int rowIndex = 2;

            // Add data
            foreach (var record in ecommerceTransactionRecords)
            {
                //System.Diagnostics.Debug.WriteLine(record.Count, "this is count......");
                //if (record.EndOfReport == "END")
                //{
                //    // Insert a blank row
                //    rowIndex++;
                //    continue;
                //}

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


                // Wrap text for multiple-line values
                worksheet.Cells.AutoFitColumns();

                rowIndex++;
            }
        }
    }
}
