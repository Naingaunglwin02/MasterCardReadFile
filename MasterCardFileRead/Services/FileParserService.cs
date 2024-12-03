using MasterCardFileRead.Models;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using static System.Runtime.InteropServices.JavaScript.JSType;

public class FileParserService
{
    public List<PresModel> ParseFile(string filePath)
    {
        var headerRecords = new List<PresModel>();
        string recordType = null, distributionID = null, cycle = null, fileId = null, date = null, memberID = null, pageNo = null, acceptanceBrand = null, transFunction = null, endOfReport = null;
        List<string> irdValues = new List<string>(), count = new List<string>(), reconAmount = new List<string>(), transferFee = new List<string>(), code = new List<string>(), proc = new List<string>();

        using (var reader = new StreamReader(filePath))
        {
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                // Parse header information
                if (line.StartsWith("1IP"))
                {
                    if (!string.IsNullOrEmpty(memberID) && !string.IsNullOrEmpty(pageNo))
                    {
                        // Finalize the current record when a new section starts
                        AddRecord();
                    }
                    recordType = line.Substring(0, 1).Trim();
                    distributionID = line.Substring(1, 16).Trim();

                    int dateStart = line.IndexOf("1IP") + "1IP".Length;
                    var parts = line.Substring(dateStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length > 0)
                    {
                        date = parts[5];
                    }

                }

                if (line.Contains("MEMBER ID:"))
                {
                    int memberStart = line.IndexOf("MEMBER ID:") + "MEMBER ID:".Length;
                    memberID = line.Substring(memberStart).Trim().Split(' ')[0];
                }

                if (line.Contains("PAGE NO:"))
                {
                    int pageStart = line.IndexOf("PAGE NO:") + "PAGE NO:".Length;
                    pageNo = line.Substring(pageStart).Trim().Split(' ')[0];
                }

                if (line.Contains("ACCEPTANCE BRAND:"))
                {
                    int acceptanceBrandStart = line.IndexOf("ACCEPTANCE BRAND:") + "ACCEPTANCE BRAND:".Length;
                    var parts = line.Substring(acceptanceBrandStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length > 0)
                    {
                        acceptanceBrand = parts[0];
                        cycle = parts[2] + " " + parts[3];

                    }
                }

                if (line.Contains("PRES.") && !line.Contains("TOTAL"))
                {
                    transFunction = line.Substring(0, 11).Trim();
                    //System.Diagnostics.Debug.WriteLine(transFunction, "Transaction Function:");
                }

                if (line.Contains("FILE ID:"))
                {
                    int fileIdStart = line.IndexOf("FILE ID:") + "FILE ID:".Length;
                    var parts = line.Substring(fileIdStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length > 0)
                    {
                        fileId = parts[0];
                    }
                }


                if (line.Contains("PURCHASE"))
                {
                    var procCode = line.Substring(0, 8).Trim();
                    proc.Add(procCode);

                    int codeStart = line.IndexOf("PURCHASE") + "PURCHASE".Length;
                    var parts = line.Substring(codeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    //Debug.WriteLine($"Parts Array: {string.Join(", ", parts)}");

                    if (parts.Length > 0)
                    {
                        code.Add(parts[0]);
                        irdValues.Add(parts[1]);
                        count.Add(parts[2]);
                        reconAmount.Add(parts[3] + " " + parts[4]);
                        transferFee.Add(parts[6] + " " + parts[7]);
                    }
                }


                if (line.Contains("***END OF REPORT***"))
                {
                    int codeStart = line.IndexOf("***") + "***".Length;

                    var parts = line.Substring(codeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length > 0)
                    {
                        endOfReport = parts[0];
                    }
                }

            }

            // Add the last record
            if (!string.IsNullOrEmpty(memberID) && !string.IsNullOrEmpty(pageNo))
            {
                AddRecord();
            }
        }

        return headerRecords;

        void AddRecord()
        {
            headerRecords.Add(new PresModel
            {
                RecordType = recordType,
                DistributionID = distributionID,
                FileId = fileId,
                Date = date,
                MemberID = memberID,
                Cycle = cycle,
                PageNo = pageNo,
                AcceptanceBrand = acceptanceBrand,
                TranscFunction = transFunction,
                Code = new List<string>(code),
                Ird = new List<string>(irdValues),
                Count = new List<string>(count),
                ReconAmount = new List<string>(reconAmount),
                TransferFee = new List<string>(transferFee),

            });

            // Reset the lists for the next record
            recordType = distributionID = memberID = pageNo = acceptanceBrand = transFunction = endOfReport = date = cycle = null;
            irdValues.Clear();
            count.Clear();
            reconAmount.Clear();
            transferFee.Clear();
            code.Clear();
        }
    }

    public List<FeeModel> ParseFileFee(string filePath)
    {
        var feeRecords = new List<FeeModel>();
        string recordType = null, distributionID = null, cycle = null, fileId = null, date = null, memberID = null, pageNo = null, acceptanceBrand = null, transFunction = null, endOfReport = null;
        List<string> irdValues = new List<string>(), count = new List<string>(), reconAmount = new List<string>(), transferFee = new List<string>(), code = new List<string>(), proc = new List<string>();

        using (var reader = new StreamReader(filePath))
        {
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                // Parse header information
                if (line.StartsWith("1IP"))
                {
                    if (!string.IsNullOrEmpty(memberID) && !string.IsNullOrEmpty(pageNo))
                    {
                        // Finalize the current record when a new section starts
                        AddRecord();
                    }
                    recordType = line.Substring(0, 1).Trim();
                    distributionID = line.Substring(1, 16).Trim();

                    int dateStart = line.IndexOf("1IP") + "1IP".Length;
                    var parts = line.Substring(dateStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length > 0)
                    {
                        date = parts[5];
                    }

                }

                if (line.Contains("MEMBER ID:"))
                {
                    int memberStart = line.IndexOf("MEMBER ID:") + "MEMBER ID:".Length;
                    memberID = line.Substring(memberStart).Trim().Split(' ')[0];
                }

                if (line.Contains("PAGE NO:"))
                {
                    int pageStart = line.IndexOf("PAGE NO:") + "PAGE NO:".Length;
                    pageNo = line.Substring(pageStart).Trim().Split(' ')[0];
                }

                if (line.Contains("ACCEPTANCE BRAND:"))
                {
                    int acceptanceBrandStart = line.IndexOf("ACCEPTANCE BRAND:") + "ACCEPTANCE BRAND:".Length;
                    var parts = line.Substring(acceptanceBrandStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length > 0)
                    {
                        acceptanceBrand = parts[0];
                        cycle = parts[2] + " " + parts[3];

                    }
                }

                if (line.Contains("FEE COLL-CSG") && !line.Contains("TOTAL"))
                {
                    transFunction = line.Substring(0, 13).Trim();
                    //System.Diagnostics.Debug.WriteLine(transFunction, "Transaction Fee:");
                }

                if (line.Contains("FILE ID:"))
                {
                    int fileIdStart = line.IndexOf("FILE ID:") + "FILE ID:".Length;
                    var parts = line.Substring(fileIdStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length > 0)
                    {
                        fileId = parts[0];
                    }
                }

                if (line.Contains("FEE COL CR"))
                {
                    //var procCode = line.Substring(0, 8).Trim();
                    //proc.Add(procCode);

                    int codeStart = line.IndexOf("FEE COL CR") + "FEE COL CR".Length;
                    var parts = line.Substring(codeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    //Debug.WriteLine($"Parts Array: {string.Join(", ", parts)}");

                    if (parts.Length > 0)
                    {
                        code.Add(parts[0]);
                        //irdValues.Add(parts[1]);
                        count.Add(parts[1]);
                        reconAmount.Add(parts[2] + " " + parts[3]);
                        transferFee.Add(parts[5] + " " + parts[6]);
                    }
                }


                if (line.Contains("***END OF REPORT***"))
                {
                    int codeStart = line.IndexOf("***") + "***".Length;

                    var parts = line.Substring(codeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length > 0)
                    {
                        //System.Diagnostics.Debug.WriteLine(string.Join(", ", parts), "This is the data after the end of the report...");
                        endOfReport = parts[0];
                    }
                }

            }

            // Add the last record
            if (!string.IsNullOrEmpty(memberID) && !string.IsNullOrEmpty(pageNo))
            {
                AddRecord();
            }
        }

        return feeRecords;

        void AddRecord()
        {
            //Debug.WriteLine($"Adding Record: MemberID={memberID}, PageNo={pageNo}, ReconAmounts={string.Join(", ", reconAmount)}");

            feeRecords.Add(new FeeModel
            {
                RecordType = recordType,
                DistributionID = distributionID,
                FileId = fileId,
                Date = date,
                MemberID = memberID,
                Cycle = cycle,
                PageNo = pageNo,
                AcceptanceBrand = acceptanceBrand,
                TranscFunction = transFunction,
                Code = new List<string>(code),
                Ird = new List<string>(irdValues),
                Count = new List<string>(count),
                ReconAmount = new List<string>(reconAmount),
                TransferFee = new List<string>(transferFee),

            });

            // Reset the lists for the next record
            recordType = distributionID = memberID = pageNo = acceptanceBrand = transFunction = endOfReport = date = cycle = null;
            irdValues.Clear();
            count.Clear();
            reconAmount.Clear();
            transferFee.Clear();
            code.Clear();
        }
    }


    public class ExcelExporter
    {
        private void AddHeaders(ExcelWorksheet worksheet, string[] headers)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
            }

            // Make the header row bold
            worksheet.Cells[1, 1, 1, headers.Length].Style.Font.Bold = true;

            for (int i = 1; i <= headers.Length; i++)
            {
                worksheet.Column(i).Width = 30;
            }
        }
        public void GenerateExcelFile(List<PresModel> headerRecords, List<FeeModel> feeRecords, string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var headerSheetRecords = headerRecords
                    .Where(record => record.TranscFunction?.Contains("FIRST PRES", StringComparison.OrdinalIgnoreCase) == true)
                    .ToList();

                var summarySheetRecords = feeRecords
                    .Where(record => record.TranscFunction?.Contains("FEE COLL-CSG", StringComparison.OrdinalIgnoreCase) == true)
                    .ToList();

                // Add Header Records sheet
                var sheet1 = package.Workbook.Worksheets.Add("Header Records");
                AddDataToSheet(sheet1, headerSheetRecords);

                // Add Summary sheet
                var sheet2 = package.Workbook.Worksheets.Add("Summary");
                if (summarySheetRecords != null && summarySheetRecords.Any())
                {
                    AddSummaryDataToSheet(sheet2, summarySheetRecords);
                }
                else
                {
                    // Optionally add a message or leave it blank if no data is available
                    sheet2.Cells[1, 1].Value = "No data available for summary.";
                }
                //AddSummaryDataToSheet(sheet2, feeRecords);

                FileInfo fileInfo = new FileInfo(filePath);
                package.SaveAs(fileInfo);
            }


        }
        private void AddDataToSheet(ExcelWorksheet worksheet, List<PresModel> headerRecords)
        {
            string[] headers = new string[]
            {
                "Record Type",
                "Distribution ID",
                "Date",
                "File ID",
                "Member ID",
                "Cycle",
                "Page No",
                "Acceptance Brand",
                "Transaction Function",
                "Code",
                "IRD Values",
                "Count",
                "Recon Amount",
                "Transfer Fee"
            };

            AddHeaders(worksheet, headers);

            int rowIndex = 2;

            // Add data
            foreach (var record in headerRecords)
            {
                if (record.EndOfReport == "END")
                {
                    // Insert a blank row
                    rowIndex++;
                    continue;
                }

                worksheet.Cells[rowIndex, 1].Value = record.RecordType;
                worksheet.Cells[rowIndex, 2].Value = record.DistributionID;
                worksheet.Cells[rowIndex, 3].Value = record.Date;
                worksheet.Cells[rowIndex, 4].Value = record.FileId;
                worksheet.Cells[rowIndex, 5].Value = record.MemberID;
                worksheet.Cells[rowIndex, 6].Value = record.Cycle;
                worksheet.Cells[rowIndex, 7].Value = record.PageNo;
                worksheet.Cells[rowIndex, 8].Value = record.AcceptanceBrand;
                worksheet.Cells[rowIndex, 9].Value = record.TranscFunction;
                worksheet.Cells[rowIndex, 10].Value = string.Join("\n", record.Code);
                worksheet.Cells[rowIndex, 11].Value = string.Join("\n", record.Ird);
                worksheet.Cells[rowIndex, 12].Value = string.Join("\n", record.Count);
                worksheet.Cells[rowIndex, 13].Value = string.Join("\n", record.ReconAmount);
                worksheet.Cells[rowIndex, 14].Value = string.Join("\n", record.TransferFee);

                // Wrap text for multiple-line values
                worksheet.Cells[rowIndex, 7, rowIndex, 14].Style.WrapText = true;

                rowIndex++;
            }
        }

        private void AddSummaryDataToSheet(ExcelWorksheet worksheet, List<FeeModel> feeRecords)
        {
            string[] headers = new string[]
             {
                "Record Type",
                "Distribution ID",
                "Date",
                "File ID",
                "Member ID",
                "Cycle",
                "Page No",
                "Acceptance Brand",
                "Transaction Function",
                "Code",
                //"IRD Values",
                "Count",
                "Recon Amount",
                "Transfer Fee"
            };


            AddHeaders(worksheet, headers);

            int rowIndex = 2;

            // Add data
            foreach (var record in feeRecords)
            {
                System.Diagnostics.Debug.WriteLine(record.TranscFunction, "this is function....");
                if (record.EndOfReport == "END")
                {
                    // Insert a blank row
                    rowIndex++;
                    continue;
                }

                worksheet.Cells[rowIndex, 1].Value = record.RecordType;
                worksheet.Cells[rowIndex, 2].Value = record.DistributionID;
                worksheet.Cells[rowIndex, 3].Value = record.Date;
                worksheet.Cells[rowIndex, 4].Value = record.FileId;
                worksheet.Cells[rowIndex, 5].Value = record.MemberID;
                worksheet.Cells[rowIndex, 6].Value = record.Cycle;
                worksheet.Cells[rowIndex, 7].Value = record.PageNo;
                worksheet.Cells[rowIndex, 8].Value = record.AcceptanceBrand;
                worksheet.Cells[rowIndex, 9].Value = record.TranscFunction;
                worksheet.Cells[rowIndex, 10].Value = string.Join("\n", record.Code);
                //worksheet.Cells[rowIndex, 11].Value = string.Join("\n", record.Ird);
                worksheet.Cells[rowIndex, 11].Value = string.Join("\n", record.Count);
                worksheet.Cells[rowIndex, 12].Value = string.Join("\n", record.ReconAmount);
                worksheet.Cells[rowIndex, 13].Value = string.Join("\n", record.TransferFee);

                // Wrap text for multiple-line values
                worksheet.Cells[rowIndex, 7, rowIndex, 13].Style.WrapText = true;

                rowIndex++;
            }
        }
    }
}

