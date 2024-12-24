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
            string errorDescription = null, date = null, fileId = null, processingMode = null, mtiFunctionCode = null, sourceMessage = null, cardNumber = null, mccCode = null, rrnCode = null, authCode = null;
            string newErrorDescriptionLine = null;
            List<string> errorCode = new List<string>();
            List<string> errorDescriptionList = new List<string>();
            List<string> elementId = new List<string>();
            using (var reader = new StreamReader(filePath))
            {
                bool isDescriptionFound = false;
                string line;

                while ((line = reader.ReadLine()) != null)
                {
                    //if (line.Contains("BUSINESS SERVICE LEVEL:"))
                    //{
                    //    date = FileReadConditionService.ExtractDate(line, ref date);
                    //    System.Diagnostics.Debug.WriteLine(date, "this is date.....");
                    //}

                    if (line.Contains("PROCESSING MODE:"))
                    {
                        processingMode = FileReadConditionService.ExtractProcessingMode(line);
                    }

                    if (line.Contains("SOURCE MESSAGE #:"))
                    {
                        sourceMessage = FileReadConditionService.ExtractResourceMessage(line);
                    }

                    if (line.Contains("MTI-FUNCTION CODE:"))
                    {
                        mtiFunctionCode = FileReadConditionService.ExtractMtiFunctionCode(line);
                        //System.Diagnostics.Debug.WriteLine(mtiFunctionCode, "this is mti function code....");
                    }

                    if (line.Contains("FILE ID:"))
                    {
                        fileId = FileReadConditionService.ExtractFileID(line);
                    }

                    if (line.Contains("DESCRIPTION") || line.Contains("MESSAGE DETAILS"))
                    {
                        //System.Diagnostics.Debug.WriteLine(line, "this is description line.....");
                        isDescriptionFound = true;
                        continue;
                    }

                    if (line.Contains("D0002"))
                    {
                        cardNumber = FileReadConditionService.ExtractD0002(line);
                    }

                    if (line.Contains("D0026"))
                    {
                        mccCode = FileReadConditionService.ExtractD0026(line);  
                    }

                    if (line.Contains("D0037"))
                    {
                        rrnCode = FileReadConditionService.ExtractD0037(line);
                    }

                    if (line.Contains("D0038"))
                    {
                        authCode = FileReadConditionService.ExtractD0038(line);

                        var transaction = new RejectTransactionModel
                        {
                            //Date = date,
                            ProcessingMode = processingMode,
                            MtiFunctionCode = mtiFunctionCode,
                            FileId = fileId,
                            ErrorCode = errorCode,
                            ErrorDescription = errorDescriptionList,
                            SourceMessage = sourceMessage,
                            ElementId = elementId,
                            CardNumberD0002 = cardNumber,
                            MccCodeD0026 = mccCode,
                            RrnD0037 = rrnCode,
                            AuthCodeD0038 = authCode,
                        };

                        rejectTransactionRecords.Add(transaction);

                    }
                }
            }

            return rejectTransactionRecords;
        }

        public void AddRejectDataToSheet(ExcelWorksheet worksheet, List<RejectTransactionModel> rejectTransactionRecords)
        {
            string[] headers = new string[]
            {
                //"Date",
                "Processing Mode",
                "MTI Function Code",
                "File Id",
                "Error Code",
                "Error Description",
                "Source Message",
                "Element ID",
                "CARD NUMBER (D0002)",
                "MCC CODE (D0026)",
                "RNN (D0037)",
                "AUTH_CODE (D0038)"
            };

            FileParserService fileParserService = new FileParserService();
            fileParserService.AddHeaders(worksheet, headers, 15);

            int rowIndex = 2;

            // Add data
            foreach (var record in rejectTransactionRecords)
            {
                //worksheet.Cells[rowIndex, 1].Value = record.Date;
                worksheet.Cells[rowIndex, 1].Value = record.ProcessingMode;
                worksheet.Cells[rowIndex, 2].Value = record.MtiFunctionCode;
                worksheet.Cells[rowIndex, 3].Value = record.FileId;
                worksheet.Cells[rowIndex, 4].Value = string.Join("\n", record.ErrorCode);
                worksheet.Cells[rowIndex, 5].Value = string.Join("\n", record.ErrorDescription);
                worksheet.Cells[rowIndex, 6].Value = string.Join("\n", record.SourceMessage);
                worksheet.Cells[rowIndex, 7].Value = string.Join("\n", record.ElementId);
                worksheet.Cells[rowIndex, 8].Value = record.CardNumberD0002;
                worksheet.Cells[rowIndex, 9].Value = record.MccCodeD0026;
                worksheet.Cells[rowIndex, 10].Value = record.RrnD0037;
                worksheet.Cells[rowIndex, 11].Value = record.AuthCodeD0038;

                worksheet.Cells[rowIndex, 4, rowIndex, 7].Style.WrapText = true;
                worksheet.Cells.AutoFitColumns();

                rowIndex++;
            }
        }
    }
}
