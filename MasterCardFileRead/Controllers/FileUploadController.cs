using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Collections.Generic;
using MasterCardFileRead.Models;
using MasterCardFileRead.Services;
using System.Transactions;

[ApiController]
[Route("api/")]
public class FileUploadController : ControllerBase
{
    private readonly EcommerceTransaction _ecommerceTransaction;
    private readonly OtherTransaction _otherTransaction;
    private readonly IssuingTransaction _issuingTransaction;

    public FileUploadController(EcommerceTransaction ecommerceTransaction, OtherTransaction otherTransaction, IssuingTransaction issuingTransaction)
    {
        _ecommerceTransaction = ecommerceTransaction;
        _otherTransaction = otherTransaction;
        _issuingTransaction = issuingTransaction;
    }

    [HttpPost("file_read")]
    public IActionResult UploadFiles([FromForm] List<IFormFile> files)
    {
        if (files == null || files.Count == 0)
        {
            return BadRequest("No files were uploaded.");
        }

        var allSections = new List<TransactionModel>();
        var allSectionsFee = new List<TransactionModel>();
        var issuingTransactionSection = new List<TransactionModel>();
        var tempFiles = new List<string>();
        string excelPath = Path.Combine(Path.GetTempPath(), "temp_file.xlsx");

        try
        {
            foreach (var file in files)
            {
                if (file.Length == 0) continue;

                // Save the uploaded file to a temporary location
                var tempFilePath = Path.GetTempFileName();
                tempFiles.Add(tempFilePath);

                using (var stream = new FileStream(tempFilePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                var sections = _ecommerceTransaction.ParseFile(tempFilePath);
                var sectionsFee = _otherTransaction.ParseFileFee(tempFilePath);
                var sectionIssuingTransaction = _issuingTransaction.IssuingTransactionService(tempFilePath);

                allSections.AddRange(sections);
                allSectionsFee.AddRange(sectionsFee);
                issuingTransactionSection.AddRange(sectionIssuingTransaction);
            }

            FileParserService fileParserService = new FileParserService();
            fileParserService.GenerateExcelFile(allSections, allSectionsFee, issuingTransactionSection, excelPath);

            var bytes = System.IO.File.ReadAllBytes(excelPath);

            allSectionsFee.Clear();

            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "result.xlsx");
        }
        finally
        {
            foreach (var tempFile in tempFiles)
            {
                if (System.IO.File.Exists(tempFile))
                {
                    System.IO.File.Delete(tempFile);
                }
            }
            if (System.IO.File.Exists(excelPath))
            {
                System.IO.File.Delete(excelPath);
            }
        }
    }
}
