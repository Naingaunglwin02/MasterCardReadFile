using MasterCardFileRead.Constants;
using MasterCardFileRead.Models;
using Microsoft.AspNetCore.Http;
using System;
using System.IO;

public static class FileReadConditionService
{
    public static string ExtractDate(string line, ref string date)
    {
        // Define the keyword to search for
        const string keyword = CommonConstants.Date;
        int dateStart = line.IndexOf(keyword) + keyword.Length;

        // Ensure the keyword exists in the line
        if (dateStart > keyword.Length - 1)
        {
            // Extract substring after the keyword
            string remainingLine = line.Substring(dateStart).Trim();

            var parts = remainingLine.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var part in parts)
            {
                // Check if part is a valid date
                if (DateOnly.TryParse(part, out DateOnly parsedDate))
                {
                    date = part;
                    return parsedDate.ToString("MM/dd/yyyy");
                }
            }
        }

        date = string.Empty;
        return string.Empty;
    }

    public static string ExtractMemberID(string line)
    {
        int memberStart = line.IndexOf(CommonConstants.MemberId) + CommonConstants.MemberId.Length;
        return line.Substring(memberStart).Trim().Split(' ')[0];
    }

    public static string ExtractAcceptanceBrandCycle(string line)
    {
        int acceptanceBrandStart = line.IndexOf(CommonConstants.Cycle) + CommonConstants.Cycle.Length;
        var parts = line.Substring(acceptanceBrandStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        return parts.Length > 3 ? parts[2] + " " + parts[3] : null;
    }

    public static string ExtractFileID(string line)
    {
        int fileIdStart = line.IndexOf(CommonConstants.FileId) + CommonConstants.FileId.Length;
        var parts = line.Substring(fileIdStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length > 0 ? parts[0] : null;
    }

    public static string ExtractFileIDForOtherTransaction(string line)
    {
        int fileIdStart = line.IndexOf(CommonConstants.FileId) + CommonConstants.FileId.Length;

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
        int fileIdStart = line.IndexOf(CommonConstants.FileId) + CommonConstants.FileId.Length;
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
        int fileIdStart = line.IndexOf(CommonConstants.FileId) + CommonConstants.FileId.Length;

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
        var keywords = new[] { CommonConstants.TransactionColCr, CommonConstants.TransactionColDr, CommonConstants.TransactionPurchase, CommonConstants.TransactionCredit, CommonConstants.TransactionAtm };

        var matchingKeyword = keywords.FirstOrDefault(keyword => line.Contains(keyword));
        if (string.IsNullOrEmpty(matchingKeyword))
        {
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
        result.ReconAmount = parts[2] + " " + parts[3];
        result.ReconDCCR = parts[3];
        result.Currency = parts[4];
        result.TransferFee = parts[5] + " " + parts[6];
        result.TransferFeeDCCR = parts[6];

        if (line.Contains(CommonConstants.TransactionPurchase) || line.Contains(CommonConstants.TransactionCredit) || line.Contains(CommonConstants.TransactionAtm))
        {
            result.TransactionFunction = beforeKeyword;
            result.Proc = matchingKeyword;
            result.Code = parts[0];
            result.IrdValues = parts[1];
            result.Count = parts[2];
            result.ReconAmount = parts[3] + " " + parts[4];
            result.ReconDCCR = parts[4];
            result.Currency = parts[5];
            result.TransferFee = parts[6] + " " +  parts[7];
            result.TransferFeeDCCR = parts[7];
        }

        return result;
    }

    public static string ExtractProcessingMode(string line)
    {
        int processingModeStart = line.IndexOf(CommonConstants.ProcessingMode) + CommonConstants.ProcessingMode.Length;
        return line.Substring(processingModeStart).Trim().Split(' ')[0];
    }

    public static string ExtractMtiFunctionCode(string line)
    {
        int mtiFunctionCodeStart = line.IndexOf(CommonConstants.MtiFunctionCode) + CommonConstants.MtiFunctionCode.Length;
        return line.Substring(mtiFunctionCodeStart).Trim().Split(' ')[0];
    }

    public static string ExtractResourceMessage(string line)
    {
        int resourceMessageCodeStart = line.IndexOf(CommonConstants.SourceMessage) + CommonConstants.SourceMessage.Length;
        string resourceMessage = line.Substring(resourceMessageCodeStart).Trim().Split(' ')[0];

        // Remove only the first leading zero if it exists
        if (resourceMessage.StartsWith("0"))
        {
            resourceMessage = resourceMessage.Substring(1);
        }

        return resourceMessage;
    }

    public static string ExtractD0002(string line)
    {
        int cardNumberCodeStart = line.IndexOf(CommonConstants.CardNumber) + CommonConstants.CardNumber.Length;
        var parts = line.Substring(cardNumberCodeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length >= 1 ? parts[0] : null;
    }

    public static string ExtractD0026(string line)
    {
        int mccCodeStart = line.IndexOf(CommonConstants.MccCode) + CommonConstants.MccCode.Length;
        var parts = line.Substring(mccCodeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length >= 1 ? parts[0] : null;
    }

    public static string ExtractD0037(string line)
    {
        int rrnCodeStart = line.IndexOf(CommonConstants.RnnCode) + CommonConstants.RnnCode.Length;
        var parts = line.Substring(rrnCodeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length >= 1 ? parts[0] : null;
    }

    public static string ExtractD0038(string line)
    {
        int authCodeStart = line.IndexOf(CommonConstants.AuthCode) + CommonConstants.AuthCode.Length;
        var parts = line.Substring(authCodeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length >= 1 ? parts[0] : null;
    }
    public static string ExtractD0041(string line)
    {
        int terminalCodeStart = line.IndexOf(CommonConstants.TerminalId) + CommonConstants.TerminalId.Length;
        var parts = line.Substring(terminalCodeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length >= 1 ? parts[0] : null;
    }

    public static string ExtractD0042(string line)
    {
        int merchantIdCodeStart = line.IndexOf(CommonConstants.MerchantId) + CommonConstants.MerchantId.Length;
        var parts = line.Substring(merchantIdCodeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length >= 1 ? parts[0] : null;
    }

    public static string ExtractD0043S01(string line)
    {
        int merchantNameCodeStart = line.IndexOf(CommonConstants.MerchantName) + CommonConstants.MerchantName.Length;
        var parts = line.Substring(merchantNameCodeStart).Trim().Split(new String[] { "  " }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length >= 1 ? parts[0] : null;
    }

    public static string ExtractP0158S04(string line)
    {
        int irdCodeStart = line.IndexOf(CommonConstants.IrdValue) + CommonConstants.IrdValue.Length;
        var parts = line.Substring(irdCodeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length >= 1 ? parts[0] : null;
    }

    public static string ExtractSourceAmount(string line)
    {
        int sourceAmountCodeStart = line.IndexOf(CommonConstants.SourceAmount) + CommonConstants.SourceAmount.Length;
        var parts = line.Substring(sourceAmountCodeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length > 0 ? parts[0] : null;
    }

    public static string ExtractSourceCurrency(string line)
    {
        int sourceCurrencyCodeStart = line.IndexOf(CommonConstants.SourceCurrency) + CommonConstants.SourceCurrency.Length;
        var parts = line.Substring(sourceCurrencyCodeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length > 0 ? parts[0] : null;
    }
}
