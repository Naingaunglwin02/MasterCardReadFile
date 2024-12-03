using MasterCardFileRead.Models;
using OfficeOpenXml;
using MasterCardFileRead.Services;

namespace MasterCardFileRead.Services
{
    public class EcommerceTransaction
    {
        public List<TransactionModel> ParseFile(string filePath)
        {
            var headerRecords = new List<TransactionModel>();
            string cycle = null, fileId = null, date = null, memberID = null, transFunction = null, endOfReport = null;
            List<string> irdValues = new List<string>(), count = new List<string>(), reconAmount = new List<string>(), transferFee = new List<string>(), code = new List<string>(), proc = new List<string>();

            using (var reader = new StreamReader(filePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    // Parse header information
                    if (line.StartsWith("1IP"))
                    {
                        if (!string.IsNullOrEmpty(memberID))
                        {
                            // Finalize the current record when a new section starts
                            AddRecord();
                        }

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

                    if (line.Contains("ACCEPTANCE BRAND:"))
                    {
                        int acceptanceBrandStart = line.IndexOf("ACCEPTANCE BRAND:") + "ACCEPTANCE BRAND:".Length;
                        var parts = line.Substring(acceptanceBrandStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                        if (parts.Length > 0)
                        {
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
                if (!string.IsNullOrEmpty(memberID))
                {
                    AddRecord();
                }
            }

            return headerRecords;

            void AddRecord()
            {
                headerRecords.Add(new TransactionModel
                {
                    TranscFunction = transFunction,
                    FileId = fileId,
                    Date = date,
                    MemberID = memberID,
                    Cycle = cycle,
                    Code = new List<string>(code),
                    Ird = new List<string>(irdValues),
                    Count = new List<string>(count),
                    ReconAmount = new List<string>(reconAmount),
                    TransferFee = new List<string>(transferFee),

                });

                // Reset the lists for the next record
                memberID = transFunction = endOfReport = date = cycle = null;
                irdValues.Clear();
                count.Clear();
                reconAmount.Clear();
                transferFee.Clear();
                code.Clear();
            }
        }

        public void AddDataToSheet(ExcelWorksheet worksheet, List<TransactionModel> headerRecords)
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
            foreach (var record in headerRecords)
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
                worksheet.Cells[rowIndex, 6].Value = string.Join("\n", record.Code);
                worksheet.Cells[rowIndex, 7].Value = string.Join("\n", record.Ird);
                worksheet.Cells[rowIndex, 8].Value = string.Join("\n", record.Count);
                worksheet.Cells[rowIndex, 9].Value = string.Join("\n", record.ReconAmount);
                worksheet.Cells[rowIndex, 10].Value = string.Join("\n", record.TransferFee);

                // Wrap text for multiple-line values
                worksheet.Cells[rowIndex, 6, rowIndex, 10].Style.WrapText = true;

                rowIndex++;
            }
        }
    }
}
