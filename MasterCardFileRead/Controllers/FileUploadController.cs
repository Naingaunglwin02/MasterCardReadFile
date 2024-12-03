using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Collections.Generic;
using MasterCardFileRead.Models;

[ApiController]
[Route("api/")]
public class FileUploadController : ControllerBase
{
    private readonly FileParserService _fileParserService;
    private static List<PresModel> _uploadedRecords = new List<PresModel>();

    public FileUploadController(FileParserService fileParserService)
    {
        _fileParserService = fileParserService;
    }

    [HttpPost("file_read")]
    public IActionResult UploadFile(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("File is empty or missing.");
        }

        var tempFilePath = Path.GetTempFileName();

        try
        {
            using (var stream = new FileStream(tempFilePath, FileMode.Create))
            {
                file.CopyTo(stream);
            }

            var sections = _fileParserService.ParseFile(tempFilePath);
            var sectionsFee = _fileParserService.ParseFileFee(tempFilePath);
            string excelPath = Path.Combine(Path.GetTempPath(), "output.xlsx");

            var exporter = new FileParserService.ExcelExporter();
            exporter.GenerateExcelFile(sections, sectionsFee, excelPath);

            var bytes = System.IO.File.ReadAllBytes(excelPath);


            // Use the ParseFile method to process all lines in the file
            var headerRecords = _fileParserService.ParseFile(tempFilePath);

            //return Ok(new
            //{
            //    headerRecords,
            //    excelDownloadLink = 
            //});

            //return Ok(headerRecords);
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "sample_sheet.xlsx");
        }
        finally
        {
            System.IO.File.Delete(tempFilePath); // Clean up temporary file
        }
    }
}
