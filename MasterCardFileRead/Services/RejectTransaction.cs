using MasterCardFileRead.Models;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace MasterCardFileRead.Services
{
    public class RejectTransaction
    {
        public List<RejectTransactionModel> RejectTransactionService(string filePath)
        {
            var rejectTransactionRecords = new List<RejectTransactionModel>();
            // Variables for transaction data
            string errorDescription = null, date = null, fileId = null, processingMode = null,
                   mtiFunctionCode = null, sourceMessage = null, cardNumber = null, mccCode = null, rrnCode = null,
                   authCode = null, terminalId = null, merchantId = null, merchantName = null, ird = null, sourceAmount = null, sourceCurrency = null, newErrorDescriptionLine = null;

            // Lists for error data
            using (var reader = new StreamReader(filePath, Encoding.GetEncoding("Windows-1252")))
            {
                string line;
                bool isMessageLevelReject = false;
                //string content = reader.ReadToEnd();
                //Console.WriteLine($"Detected Encoding: {reader.CurrentEncoding.EncodingName}");

                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Contains("MESSAGE LEVEL REJECT"))
                    {
                        isMessageLevelReject = true;
                        continue;
                    }

                    if (isMessageLevelReject)
                    {
                        if (Regex.IsMatch(line, @"\d{4}/\d{2}/\d{2}"))
                        {
                            var data = line.Trim();
                            date = DateOnly.Parse(data).ToString("MM/dd/yyyy");
                            isMessageLevelReject = false;
                        }
                    }

                    if (line.Contains("SOURCE MESSAGE #:"))
                        sourceMessage = FileReadConditionService.ExtractResourceMessage(line);

                    if (line.Contains("MTI-FUNCTION CODE:"))
                        mtiFunctionCode = FileReadConditionService.ExtractMtiFunctionCode(line);

                    if (line.Contains("FILE ID:"))
                        fileId = FileReadConditionService.ExtractFileID(line);

                    if (line.Contains("D0002"))
                        cardNumber = FileReadConditionService.ExtractD0002(line);

                    if (line.Contains("D0026"))
                        mccCode = FileReadConditionService.ExtractD0026(line);

                    if (line.Contains("D0037"))
                        rrnCode = FileReadConditionService.ExtractD0037(line);

                    if (line.Contains("D0038"))
                        authCode = FileReadConditionService.ExtractD0038(line);

                    if (line.Contains("D0041"))
                        terminalId = FileReadConditionService.ExtractD0041(line);

                    if (line.Contains("D0042"))
                        merchantId = FileReadConditionService.ExtractD0042(line);

                    if (line.Contains("D0043 S01"))
                        merchantName = FileReadConditionService.ExtractD0043S01(line);

                    if (line.Contains("P0158 S04"))
                        ird = FileReadConditionService.ExtractP0158S04(line);

                    if (line.Contains("SOURCE AMOUNT:"))
                        sourceAmount = FileReadConditionService.ExtractSourceAmount(line);

                    if (line.Contains("SOURCE CURRENCY:"))
                        sourceCurrency = FileReadConditionService.ExtractSourceCurrency(line);

                    if (line.Contains("PROCESSING MODE:"))
                    {
                        processingMode = FileReadConditionService.ExtractProcessingMode(line);

                        if (!string.IsNullOrEmpty(cardNumber))
                        {
                            var transaction = new RejectTransactionModel
                            {
                                Date = date,
                                ProcessingMode = processingMode,
                                MtiFunctionCode = mtiFunctionCode,
                                FileId = fileId,
                                SourceMessage = sourceMessage,
                                CardNumberD0002 = cardNumber,
                                MccCodeD0026 = mccCode,
                                RrnD0037 = rrnCode,
                                AuthCodeD0038 = authCode,
                                TerminalIdD0041 = terminalId,
                                MerchantIdD0042 = merchantId,
                                MerchantNameD0043S01 = merchantName,
                                IrdP0158S04 = ird,
                                SourceAmount = sourceAmount,
                                SourceCurrency = sourceCurrency
                            };

                            rejectTransactionRecords.Add(transaction);

                            errorDescription = date = fileId = processingMode =
                                   mtiFunctionCode = sourceMessage = cardNumber = mccCode = rrnCode =
                                   authCode = terminalId = merchantId = merchantName = ird = sourceAmount = sourceCurrency = null;
                        }

                    }

                }
            }
            return rejectTransactionRecords;
        }

        public Dictionary<CompositeKey, ErrorDescriptionModel> RejectTransactionDescriptionService(string filePath)
        {
            string newErrorDescriptionLine = null, errorDescription = null;
            bool isDescriptionFound = false;
            bool isMTI = false;
            bool isMessageDetailFound = false;
            string date = null;
            string sourceMessage = null;
            string mtiFunctionCode = null;

            Dictionary<CompositeKey, ErrorDescriptionModel> errorDescriptionModel = new();
            string errorTemp = null;
            string elementTemp = null;

            using (var reader = new StreamReader(filePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Contains("BUSINESS SERVICE LEVEL:"))
                    {
                        date = FileReadConditionService.ExtractDate(line, ref date);
                    }
                    if (line.Contains("MTI-FUNCTION CODE: 1240-200"))
                    {
                        isMTI = true;
                        mtiFunctionCode = FileReadConditionService.ExtractMtiFunctionCode(line);
                        continue;
                    }
                    // Check for DESCRIPTION section start
                    if (isMTI && line.Contains("DESCRIPTION"))
                    {
                        isDescriptionFound = true;
                        continue; // Skip the "DESCRIPTION" line
                    }

                    // Check for MESSAGE DETAILS section start
                    if (line.Contains("MESSAGE DETAILS"))
                    {

                        isDescriptionFound = false;
                        isMTI = false;
                        continue;
                    }

                    // Process lines between DESCRIPTION and MESSAGE DETAILS
                    if (isDescriptionFound && !string.IsNullOrEmpty(mtiFunctionCode) && !string.IsNullOrEmpty(line))
                    {
                        string[] parts = line.Split(new[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length == 1 || parts.Length == 2)
                        {
                            newErrorDescriptionLine = newErrorDescriptionLine + parts[0];
                        }
                        else
                        {
                            newErrorDescriptionLine = null;
                            errorTemp = parts[0];
                            errorDescription = parts[1];
                            sourceMessage = parts[2];
                            elementTemp = parts[3];
                        }

                        if (!string.IsNullOrEmpty(newErrorDescriptionLine))
                        {
                            errorDescription = string.Join(" ", errorDescription, newErrorDescriptionLine);

                            if (!string.IsNullOrEmpty(sourceMessage))
                            {
                                var compositeKey = new CompositeKey
                                {
                                    SourceMessage = sourceMessage.Trim(),
                                    Date = date
                                };

                                if (errorDescriptionModel.TryGetValue(compositeKey, out ErrorDescriptionModel temp))
                                {
                                    errorDescriptionModel[compositeKey].ErrorCode.Add(errorTemp);
                                    errorDescriptionModel[compositeKey].Description.Add(errorDescription + "\n");
                                    errorDescriptionModel[compositeKey].ElementId.Add(elementTemp);
                                }
                                else
                                {
                                    var tempModel = new ErrorDescriptionModel();
                                    tempModel.ErrorCode.Add(errorTemp);
                                    tempModel.Description.Add(errorDescription + "\n");
                                    tempModel.ElementId.Add(elementTemp);
                                    errorDescriptionModel[compositeKey] = tempModel;
                                }
                                sourceMessage = null;
                            }
                        }
                    }
                }

            }

            return errorDescriptionModel;
        }

        public void AddRejectDataToSheet(ExcelWorksheet worksheet, List<RejectTransactionModel> rejectTransactionRecords, Dictionary<CompositeKey, ErrorDescriptionModel> rejectTransactionDescriptionRecords)
        {

            string[] headers = new string[]
            {
                "DATE",
                "PROCESSING MODE",
                "MTI-FUNCTION CODE",
                "FILE ID",
                "ERROR CODE",
                "ERROR DESCRIPTION",
                "SOURCE MESSAGE",
                "ELEMENT ID",
                "CARD NUMBER (D0002)",
                "MCC CODE (D0026)",
                "RNN (D0037)",
                "AUTH_CODE (D0038)",
                "THERMINAL ID (D0041)",
                "MERCHANT ID (D0042)",
                "MERCHANT NAME (D0043 S01)",
                "IRD (P0158 S04)",
                "SOURCE AMOUNT",
                "SOURCE CURRENCY",
            };

            FileParserService fileParserService = new FileParserService();
            fileParserService.AddHeaders(worksheet, headers, 15);

            int rowIndex = 2;
            string previousDate = null;
            string errorDescription = null;
            double totalSourceAmount = 0;
            double grandTotalSourceAmount = 0;

            foreach (var record in rejectTransactionRecords)
            {
                var compositeKey = new CompositeKey
                {
                    SourceMessage = record.SourceMessage.Trim(),
                    Date = record.Date
                };
                var matchingRecords = rejectTransactionDescriptionRecords[compositeKey];

                if (previousDate != null && record.Date != previousDate)
                {
                    // Add total row for the previous date
                    TotalTransactions.AddSubTotalOfRejectRow(worksheet, rowIndex, "Total", headers, totalSourceAmount);
                    rowIndex += 2;
                    TotalTransactions.ResetSubtotalRejectVaribles(ref totalSourceAmount);

                    TotalTransactions.AddGrandTotalOfRejectRow(worksheet, rowIndex, "Grand Total", headers, grandTotalSourceAmount);
                    rowIndex += 2;
                    TotalTransactions.ResetGrandtoalRejectVariables(ref grandTotalSourceAmount);
                }

                previousDate = record.Date;

                // Populate the worksheet with the filtered data
                worksheet.Cells[rowIndex, 1].Value = record.Date;
                worksheet.Cells[rowIndex, 2].Value = record.ProcessingMode;
                worksheet.Cells[rowIndex, 3].Value = record.MtiFunctionCode;
                worksheet.Cells[rowIndex, 4].Value = record.FileId;
                worksheet.Cells[rowIndex, 5].Value = string.Join("\n", matchingRecords.ErrorCode);
                worksheet.Cells[rowIndex, 6].Value = string.Join("\n", matchingRecords.Description);
                worksheet.Cells[rowIndex, 7].Value = record.SourceMessage;
                worksheet.Cells[rowIndex, 8].Value = string.Join("\n", matchingRecords.ElementId);
                worksheet.Cells[rowIndex, 9].Value = record.CardNumberD0002;
                worksheet.Cells[rowIndex, 10].Value = record.MccCodeD0026;
                worksheet.Cells[rowIndex, 11].Value = record.RrnD0037;
                worksheet.Cells[rowIndex, 12].Value = record.AuthCodeD0038;
                worksheet.Cells[rowIndex, 13].Value = record.TerminalIdD0041;
                worksheet.Cells[rowIndex, 14].Value = record.MerchantIdD0042;
                worksheet.Cells[rowIndex, 15].Value = record.MerchantNameD0043S01;
                worksheet.Cells[rowIndex, 16].Value = record.IrdP0158S04;
                worksheet.Cells[rowIndex, 17].Value = record.SourceAmount;
                worksheet.Cells[rowIndex, 18].Value = record.SourceCurrency;

                totalSourceAmount += Convert.ToDouble(record.SourceAmount);
                grandTotalSourceAmount += Convert.ToDouble(record.SourceAmount);

                worksheet.Cells[rowIndex, 5, rowIndex, 8].Style.WrapText = true;
                worksheet.Cells[rowIndex, 1, rowIndex, 18].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                worksheet.Cells[rowIndex, 1, rowIndex, 18].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                worksheet.Cells.AutoFitColumns();

                rowIndex++;

            }
            if (previousDate != null)
            {
                // Add final total row
                TotalTransactions.AddSubTotalOfRejectRow(worksheet, rowIndex, "Total", headers, totalSourceAmount);
                rowIndex += 2;

                TotalTransactions.AddGrandTotalOfRejectRow(worksheet, rowIndex, "Grand Total", headers, grandTotalSourceAmount);
                rowIndex += 2;

            }

        }
    }
}
