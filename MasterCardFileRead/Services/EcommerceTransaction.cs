using MasterCardFileRead.Models;
using OfficeOpenXml;
using MasterCardFileRead.Services;

namespace MasterCardFileRead.Services
{
    public class EcommerceTransaction
    {
        public List<TransactionModel> ParseFile(string filePath)
        {
            var ecommerceTransactionRecords = new List<TransactionModel>();

            using (var reader = new StreamReader(filePath))
            {
                string line;
                string date = null, memberID = null, cycle = null, fileId = null;

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
                        fileId = FileReadConditionService.ExtractFileIDEven(line);

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

                        ecommerceTransactionRecords.Add(transaction);
                    }
                }
            }

            return ecommerceTransactionRecords;
        }

        public void AddDataToSheet(ExcelWorksheet worksheet, List<TransactionModel> ecommerceTransactionRecords)
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

            string previousDate = null;
            int totalCount = 0;
            double totalRecon = 0;
            double totalTransFee = 0;

            string totalCr = "";
            string totalDr = "";

            // Add data
            foreach (var record in ecommerceTransactionRecords)
            {
                if (previousDate != null && record.Date != previousDate)
                {
                    //rowIndex++;
                    //worksheet.Cells[rowIndex, 3].Value = "Total";
                    worksheet.Cells[rowIndex, 1, rowIndex, 8].Merge = true;      
                    worksheet.Cells[rowIndex, 1].Value = "Total";

                    worksheet.Cells[rowIndex, 9].Value = totalCount;
                    worksheet.Cells[rowIndex, 10].Value = totalRecon;
                    worksheet.Cells[rowIndex, 11].Value = record.ReconDCCR;

                    worksheet.Cells[rowIndex, 13].Value = totalTransFee;
                    worksheet.Cells[rowIndex, 14].Value = totalDr;

                    using (var range = worksheet.Cells[rowIndex, 1, rowIndex, 14])
                    {
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        range.Style.Font.Bold = true;

                        // Center-align the "Total" text in the merged cells (columns 1 to 8)
                        worksheet.Cells[rowIndex, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[rowIndex, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        // Write the Total Count in column 9 and left-align it
                        worksheet.Cells[rowIndex, 9].Value = totalCount;
                        worksheet.Cells[rowIndex, 9].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        worksheet.Cells[rowIndex, 10].Value = totalRecon;
                        worksheet.Cells[rowIndex, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        worksheet.Cells[rowIndex, 13].Value = totalTransFee;
                        worksheet.Cells[rowIndex, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        worksheet.Cells[rowIndex, 11].Value = totalCr;
                        worksheet.Cells[rowIndex, 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        worksheet.Cells[rowIndex, 14].Value = totalDr;
                        worksheet.Cells[rowIndex, 14].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    }

                    totalCount = 0;
                    totalRecon = 0;
                    totalTransFee = 0;
                    totalCr = "";
                    totalDr = "";


                    rowIndex += 2;
                }

                System.Diagnostics.Debug.WriteLine("Hello");

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


                totalCount += Int32.Parse(record.Count);
                //totalRecon += Int32.Parse(record.ReconAmount);
                totalRecon += Convert.ToDouble(record.ReconAmount);
                totalTransFee += Convert.ToDouble(record.TransferFee);
                totalCr = record.ReconDCCR;
                totalDr = record.TransferFeeDCCR;

                worksheet.Cells.AutoFitColumns();

                previousDate = record.Date;
                rowIndex++;
            }

            if (previousDate != null)
            {
           
                worksheet.Cells[rowIndex, 1, rowIndex, 8].Merge = true;
                worksheet.Cells[rowIndex, 1].Value = "Total";

                worksheet.Cells[rowIndex, 9].Value = totalCount;

                // Apply gray background and bold font to the final "Total" row
                using (var range = worksheet.Cells[rowIndex, 1, rowIndex, 14])
                {
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    range.Style.Font.Bold = true;

                    // Center-align the "Total" text in the merged cells (columns 1 to 8)
                    worksheet.Cells[rowIndex, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells[rowIndex, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    // Write the Total Count in column 9 and left-align it
                    worksheet.Cells[rowIndex, 9].Value = totalCount;
                    worksheet.Cells[rowIndex, 9].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    worksheet.Cells[rowIndex, 10].Value = totalRecon;
                    worksheet.Cells[rowIndex, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    worksheet.Cells[rowIndex, 13].Value = totalTransFee;
                    worksheet.Cells[rowIndex, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    worksheet.Cells[rowIndex, 11].Value = totalCr;
                    worksheet.Cells[rowIndex, 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    worksheet.Cells[rowIndex, 14].Value = totalDr;
                    worksheet.Cells[rowIndex, 14].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;





                }

            }

        }
    }
}
