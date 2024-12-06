﻿using MasterCardFileRead.Models;

public static class FileReadConditionService
{
    public static string ExtractDate(string line, ref string date)
    {
        int dateStart = line.IndexOf("1IP") + "1IP".Length;
        var parts = line.Substring(dateStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length > 5)
        {
            date = parts[5];
        }
        return date;
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

    public static TransactionResult ProcessOtherTransaction(string line)
    {
        var result = new TransactionResult();
        var keywords = new[] { "FEE COL CR", "FEE COL DR" };

        var matchingKeyword = keywords.FirstOrDefault(line.Contains);
        int keywordStart = line.IndexOf(matchingKeyword);
        int codeStart = keywordStart + matchingKeyword.Length;

        string beforeKeyword = line.Substring(0, keywordStart).Trim();

        var parts = line.Substring(codeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        result.TransactionFunction = beforeKeyword;
        result.Code = parts[0];
        result.Count = parts[1];
        result.ReconAmount = parts[2] + " " + parts[3];
        result.TransferFee = parts[5] + " " + parts[6];

        return result;
    }

    public static TransactionResult ProcessEcommerceTransaction(string line)
    {
        var result = new TransactionResult();
        var keywords = new[] { "PURCHASE" };

        var matchingKeyword = keywords.FirstOrDefault(line.Contains);
        int keywordStart = line.IndexOf(matchingKeyword);
        int codeStart = keywordStart + matchingKeyword.Length;

        string beforeKeyword = line.Substring(0, keywordStart).Trim();

        var parts = line.Substring(codeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        result.TransactionFunction = beforeKeyword;
        result.Code = parts[0];
        result.IrdValues = parts[1];
        result.Count = parts[2];
        result.ReconAmount = parts[3] + " " + parts[4];
        result.TransferFee = parts[6] + " " + parts[7];

        return result;
    }

    public static string ExtractEndOfReport(string line)
    {
        int codeStart = line.IndexOf("***") + "***".Length;
        var parts = line.Substring(codeStart).Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return parts.Length > 0 ? parts[0] : null;
    }
}