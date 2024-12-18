using MasterCardFileRead.Models;
using OfficeOpenXml;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace MasterCardFileRead.Services
{
    public class RejectTransaction
    {
        public List<RejectTransactionModel> RejectTransactionService(string filePath)
        {
            var rejectTransactionRecords = new List<RejectTransactionModel>();
            string errorDescription = null, errorCode = null, sourceMessage = null, elementId = null, date = null, fileId = null, processingMode = null, mtiFunctionCode = null;
            string newErrorDescriptionLine = null;
            using (var reader = new StreamReader(filePath))
            {
                bool isDescriptionFound = false;
                string line;

                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith("1IP"))
                    {
                        date = FileReadConditionService.ExtractDate(line, ref date);
                    }

                    if (line.Contains("PROCESSING MODE:"))
                    {
                        processingMode = FileReadConditionService.ExtractProcessingMode(line);
                    }

                    if (line.Contains("MTI-FUNCTION CODE:"))
                    {
                        mtiFunctionCode = FileReadConditionService.ExtractMtiFunctionCode(line);
                    }

                    if (line.Contains("FILE ID:"))
                    {
                        fileId = FileReadConditionService.ExtractFileIDEven(line);
                    }

                    if (line.Contains("DESCRIPTION"))
                    {
                        isDescriptionFound = true;
                        continue;
                    }

                    if (isDescriptionFound && !string.IsNullOrWhiteSpace(line))
                    {
                        if (line.Contains("MESSAGE DETAILS"))
                        {
                            break;
                        }
                        string[] parts = line.Split(new[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length == 1)
                        {
                            newErrorDescriptionLine = parts[0];
                        }
                        else
                        {
                            newErrorDescriptionLine = null;
                            errorCode = parts[0];
                            errorDescription = parts[1];
                            sourceMessage = parts[2];
                            elementId = parts[3];
                        }

                        if (!string.IsNullOrEmpty(newErrorDescriptionLine))
                        {
                            errorDescription = string.Join(" ", errorDescription, newErrorDescriptionLine);
                            var transaction = new RejectTransactionModel
                            {
                                Date = date,
                                ProcessingMode = processingMode,
                                MtiFunctionCode = mtiFunctionCode,
                                FileId = fileId,
                                ErrorCode = errorCode,
                                ErrorDescription = errorDescription,
                                SourceMessage = sourceMessage,
                                ElementId = elementId
                            };

                            rejectTransactionRecords.Add(transaction);
                        }
                    }
                }
            }

            return rejectTransactionRecords;
        }

        public void AddRejectDataToSheet(ExcelWorksheet worksheet, List<RejectTransactionModel> rejectTransactionRecords)
        {
            string[] headers = new string[]
            {
                "Date",
                "Processing Mode",
                "MTI Function Code",
                "File Id",
                "Error Code",
                "Error Description",
                "Source Message",
                "Element ID"
            };

            FileParserService fileParserService = new FileParserService();
            fileParserService.AddHeaders(worksheet, headers);

            int rowIndex = 2;

            // Add data
            foreach (var record in rejectTransactionRecords)
            {
                //System.Diagnostics.Debug.WriteLine(record, "this is record....");
                //if (record.EndOfReport == "END")
                //{
                //    // Insert a blank row
                //    rowIndex++;
                //    continue;
                //}
                worksheet.Cells[rowIndex, 1].Value = record.Date;
                worksheet.Cells[rowIndex, 2].Value = record.ProcessingMode;
                worksheet.Cells[rowIndex, 3].Value = record.MtiFunctionCode;
                worksheet.Cells[rowIndex, 4].Value = record.FileId;
                worksheet.Cells[rowIndex, 5].Value = string.Join("\n", record.ErrorCode);
                worksheet.Cells[rowIndex, 6].Value = string.Join("\n", record.ErrorDescription);
                worksheet.Cells[rowIndex, 7].Value = string.Join("\n", record.SourceMessage);
                worksheet.Cells[rowIndex, 8].Value = string.Join("\n", record.ElementId);

                worksheet.Cells[rowIndex, 5, rowIndex, 8].Style.WrapText = true;
                worksheet.Cells.AutoFitColumns();

                rowIndex++;
            }
        }
    }
}
