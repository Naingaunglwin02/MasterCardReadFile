using MasterCardFileRead.Models;
using Microsoft.AspNetCore.Http;
using System;

public static class FileReadConditionService
{
    public static string ExtractDate(string line, ref string date)
    {
        int dateStart = line.IndexOf("BUSINESS SERVICE LEVEL:") + "BUSINESS SERVICE LEVEL:".Length;
        var parts = line.Substring(dateStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length > 0)
        {
            date = parts[1];
        }
        var dateOnlyString = DateOnly.Parse(date).ToString("MM/dd/yyyy");

        return dateOnlyString;
    }

    public static string ExtractMemberID(string line)
    {
        int memberStart = line.IndexOf("MEMBER ID:") + "MEMBER ID:".Length;
        return line.Substring(memberStart).Trim().Split(' ')[0];
    }

    public static string ExtractAcceptanceBrandCycle(string line)
    {
        int acceptanceBrandStart = line.IndexOf("ACCEPTANCE BRAND:") + "ACCEPTANCE BRAND:".Length;
        var parts = line.Substring(acceptanceBrandStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        return parts.Length > 3 ? parts[2] + " " + parts[3] : null;
    }

    public static string ExtractFileID(string line)
    {
        int fileIdStart = line.IndexOf("FILE ID:") + "FILE ID:".Length;
        var parts = line.Substring(fileIdStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length > 0 ? parts[0] : null;
    }

    public static string ExtractFileIDForOtherTransaction(string line)
    {
        int fileIdStart = line.IndexOf("FILE ID:") + "FILE ID:".Length;

        string fileIdPart = line.Substring(fileIdStart).Trim();

        var parts = fileIdPart.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        if (parts.Length > 0)
        {
            string candidateFileId = parts[0];

            var segments = candidateFileId.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
           
            if (segments.Length > 0)
            {
                string lastSegment = segments[^1];
                if (int.TryParse(lastSegment, out int lastValue) && lastValue > 9)
                {
                    return candidateFileId;
                }
            }
        }
        return null;
    }

    public static string ExtractFileIDOdd(string line)
    {
        int fileIdStart = line.IndexOf("FILE ID:") + "FILE ID:".Length;
        var parts = line.Substring(fileIdStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        string lastPart = parts[^1]; // Gets the last part after the last '/'
       
        if (parts.Length > 0)
        {

            if (!string.IsNullOrEmpty(lastPart) && char.IsDigit(lastPart[^1]))
            {
                int lastDigit = int.Parse(lastPart[^1].ToString());

                string candidateFileId = parts[0];
                var segments = candidateFileId.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
               
                if (segments.Length > 0)
                {
                    string lastSegment = segments[^1];
                    if (lastDigit % 2 != 0 && int.TryParse(lastSegment, out int lastValue) && lastValue < 9)
                    {
                        return candidateFileId;
                    }
                }
            }
        }
      
        return null; 
    }

    public static string ExtractFileIDEven(string line)
    {
        // Locate the starting position of "FILE ID:"
        int fileIdStart = line.IndexOf("FILE ID:") + "FILE ID:".Length;

        // Extract the substring after "FILE ID:" and split by spaces
        var parts = line.Substring(fileIdStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        // Take the last part
        string lastPart = parts[^1]; // Access the last element of the array

        // Check if the last character of the last part is a digit
        if (parts.Length > 0)
        {

            if (!string.IsNullOrEmpty(lastPart) && char.IsDigit(lastPart[^1]))
            {
                int lastDigit = int.Parse(lastPart[^1].ToString());

                string candidateFileId = parts[0];
                var segments = candidateFileId.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

                if (segments.Length > 0)
                {
                    string lastSegment = segments[^1];
                    if (lastDigit % 2 == 0 && int.TryParse(lastSegment, out int lastValue) && lastValue < 9)
                    {
                        return candidateFileId;
                    }
                }
            }
        }

        // Return the first part of the split (typically the file ID)
        return null;
    }

    public static TransactionResult ProcessIssuingTransaction(string line)
    {
        var result = new TransactionResult();
        var keywords = new[] { "FEE COL CR", "FEE COL DR", "PURCHASE", "CREDIT" };

        var matchingKeyword = keywords.FirstOrDefault(keyword => line.Contains(keyword));
        if (string.IsNullOrEmpty(matchingKeyword))
        {
            //System.Diagnostics.Debug.WriteLine("No matching keyword found in line.", "Keyword Check");
            return null;
        }
        int keywordStart = line.IndexOf(matchingKeyword);
        int codeStart = keywordStart + matchingKeyword.Length;

        string beforeKeyword = line.Substring(0, keywordStart).Trim();

        var parts = line.Substring(codeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        result.TransactionFunction = beforeKeyword;
        result.Proc = matchingKeyword;
        result.Code = parts[0];
        result.Count = parts[1];
        result.ReconAmount = parts[2];
        result.ReconDCCR = parts[3];
        result.Currency = parts[4];
        result.TransferFee = parts[5];
        result.TransferFeeDCCR = parts[6];

        if (line.Contains("PURCHASE") || line.Contains("CREDIT"))
        {
            result.TransactionFunction = beforeKeyword;
            result.Proc = matchingKeyword;
            result.Code = parts[0];
            result.IrdValues = parts[1];
            result.Count = parts[2];
            result.ReconAmount = parts[3];
            result.ReconDCCR = parts[4];
            result.Currency = parts[5];
            result.TransferFee = parts[6];
            result.TransferFeeDCCR = parts[7];
        }

        return result;
    }

    public static string ExtractProcessingMode(string line)
    {
        int processingModeStart = line.IndexOf("PROCESSING MODE:") + "PROCESSING MODE:".Length;
        return line.Substring(processingModeStart).Trim().Split(' ')[0];
    }

    public static string ExtractMtiFunctionCode(string line)
    {
        int mtiFunctionCodeStart = line.IndexOf("MTI-FUNCTION CODE:") + "MTI-FUNCTION CODE:".Length;
        return line.Substring(mtiFunctionCodeStart).Trim().Split(' ')[0];
    }

    public static string ExtractEndOfReport(string line)
    {
        if(line.Contains("***END OF REPORT***"))
        {
            System.Diagnostics.Debug.WriteLine(line, "this is line.....");
            int codeStart = line.IndexOf("***") + "***".Length;
            //System.Diagnostics.Debug.WriteLine(codeStart, "this is code Start.......");
            var parts = line.Substring(codeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 0)
            {
                return parts[0];
            }
        }

        return null;
    }
}
