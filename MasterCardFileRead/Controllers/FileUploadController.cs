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
    private readonly RejectTransaction _rejectTransaction;
    private readonly PosTransaction _posTransaction;


    public FileUploadController(EcommerceTransaction ecommerceTransaction, OtherTransaction otherTransaction, IssuingTransaction issuingTransaction, RejectTransaction rejectTransaction, PosTransaction posTransaction)
    {
        _ecommerceTransaction = ecommerceTransaction;
        _otherTransaction = otherTransaction;
        _issuingTransaction = issuingTransaction;
        _rejectTransaction = rejectTransaction;
        _posTransaction = posTransaction;
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
        var rejectTransactionSection = new List<RejectTransactionModel>();
        var rejectTransactionDescriptionSection = new Dictionary<string, ErrorDescriptionModel>();
        var allSectionsPos = new List<TransactionModel>();

        var tempFiles = new List<string>();
        string excelPath = Path.Combine(Path.GetTempPath(), "temp_file.xlsx");

        Dictionary<string, ErrorDescriptionModel> sectionDescriptionRejectTransaction = new();

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
                var sectionsPos = _posTransaction.ParseFilePos(tempFilePath);

                var sectionIssuingTransaction = _issuingTransaction.IssuingTransactionService(tempFilePath);
                var sectionRejectTransaction = _rejectTransaction.RejectTransactionService(tempFilePath);

                // Populate `sectionDescriptionRejectTransaction` by merging or appending entries
                var tempDescriptionRejectTransaction = _rejectTransaction.RejectTransactionDescriptionService(tempFilePath);

                foreach (var kvp in tempDescriptionRejectTransaction)
                {
                    if (sectionDescriptionRejectTransaction.TryGetValue(kvp.Key, out var existingModel))
                    {
                        existingModel.ErrorCode.AddRange(kvp.Value.ErrorCode);
                        existingModel.Description.AddRange(kvp.Value.Description);
                        existingModel.ElementId.AddRange(kvp.Value.ElementId);
                    }
                    else
                    {
                        sectionDescriptionRejectTransaction[kvp.Key] = kvp.Value;
                    }
                }

                allSections.AddRange(sections);
                allSectionsFee.AddRange(sectionsFee);
                issuingTransactionSection.AddRange(sectionIssuingTransaction);
                rejectTransactionSection.AddRange(sectionRejectTransaction);
                allSectionsPos.AddRange(sectionsPos);
            }

            FileParserService fileParserService = new FileParserService();
            fileParserService.GenerateExcelFile(allSections, allSectionsFee, issuingTransactionSection, allSectionsPos, rejectTransactionSection, sectionDescriptionRejectTransaction, excelPath);

            var bytes = System.IO.File.ReadAllBytes(excelPath);

            allSectionsFee.Clear();
            allSectionsPos.Clear();

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
