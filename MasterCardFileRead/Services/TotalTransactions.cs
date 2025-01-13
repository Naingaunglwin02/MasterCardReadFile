using OfficeOpenXml;
using System.Reflection.PortableExecutable;

namespace MasterCardFileRead.Services
{
    public class TotalTransactions
    {
        public static void AddSubtotalRow(ExcelWorksheet worksheet, int rowIndex, string title, int totalCount, double totalRecon, double totalTransFee, string totalCr, string totalDr)
        {
            worksheet.Cells[rowIndex, 1, rowIndex, 8].Merge = true;
            worksheet.Cells[rowIndex, 1].Value = title;

            worksheet.Cells[rowIndex, 9].Value = totalCount;
            worksheet.Cells[rowIndex, 10].Value = totalRecon;
            worksheet.Cells[rowIndex, 10].Style.Numberformat.Format = "###0.00";
            worksheet.Cells[rowIndex, 11].Value = totalCr;
            worksheet.Cells[rowIndex, 13].Value = totalTransFee;
            worksheet.Cells[rowIndex, 13].Style.Numberformat.Format = "###0.00";
            worksheet.Cells[rowIndex, 14].Value = totalDr;

            ApplySubTotalRowStyle(worksheet, rowIndex);
        }

        public static void AddGrandTotalRow(ExcelWorksheet worksheet, int rowIndex, string title, int grandTotalCount, double grandTotalRecon, double grandTotalTransFee)
        {
            worksheet.Cells[rowIndex, 1, rowIndex, 8].Merge = true;
            worksheet.Cells[rowIndex, 1].Value = title;

            worksheet.Cells[rowIndex, 9].Value = grandTotalCount;
            worksheet.Cells[rowIndex, 10].Value = grandTotalRecon;
            worksheet.Cells[rowIndex, 10].Style.Numberformat.Format = "###0.00";
            worksheet.Cells[rowIndex, 13].Value = grandTotalTransFee;
            worksheet.Cells[rowIndex, 13].Style.Numberformat.Format = "###0.00";

            ApplyGrandTotalRowStyle(worksheet, rowIndex);
        }

        public static void AddSubTotalOfRejectRow(ExcelWorksheet worksheet, int rowIndex, string title, string[] headers, double subTotalSourceAmount)
        {
            worksheet.Cells[rowIndex, 1, rowIndex, headers.Length - 2].Merge = true;
            worksheet.Cells[rowIndex, 1].Value = title;

            worksheet.Cells[rowIndex, headers.Length - 1].Value = subTotalSourceAmount;
            worksheet.Cells[rowIndex, headers.Length - 1].Style.Numberformat.Format = "#,##0.00";

            ApplySubTotalRejectRowStyle(worksheet, rowIndex, headers);
        }

        public static void AddGrandTotalOfRejectRow(ExcelWorksheet worksheet, int rowIndex, string title, string[] headers, double grandTotalSourceAmount)
        {

            worksheet.Cells[rowIndex, 1, rowIndex, headers.Length - 2].Merge = true;
            worksheet.Cells[rowIndex, 1].Value = title;

            worksheet.Cells[rowIndex, headers.Length - 1].Value = grandTotalSourceAmount;
            worksheet.Cells[rowIndex, headers.Length - 1].Style.Numberformat.Format = "#,##0.00";

            ApplyGrandTotalRejectRowStyle(worksheet, rowIndex, headers);
        }

        public static void ApplySubTotalRowStyle(ExcelWorksheet worksheet, int rowIndex)
        {
            using (var range = worksheet.Cells[rowIndex, 1, rowIndex, 14])
            {
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                range.Style.Font.Bold = true;
                range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            }
            worksheet.Cells[rowIndex, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowIndex, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            rowIndex++;

        }

        public static void ApplyGrandTotalRowStyle(ExcelWorksheet worksheet, int rowIndex)
        {
            using (var range = worksheet.Cells[rowIndex, 1, rowIndex, 14])
            {
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
                range.Style.Font.Bold = true;
                range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            }
            worksheet.Cells[rowIndex, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells[rowIndex, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            rowIndex++;
        }

        public static void ApplyGrandTotalRejectRowStyle(ExcelWorksheet worksheet, int rowIndex, string[] headers)
        {
            using (var range = worksheet.Cells[rowIndex, 1, rowIndex, headers.Length])
            {
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
                range.Style.Font.Bold = true;

                worksheet.Cells[rowIndex, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[rowIndex, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                worksheet.Cells[rowIndex, headers.Length - 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            }
        }

        public static void ApplySubTotalRejectRowStyle(ExcelWorksheet worksheet, int rowIndex, string[] headers)
        {
            using (var range = worksheet.Cells[rowIndex, 1, rowIndex, headers.Length])
            {
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                range.Style.Font.Bold = true;

                worksheet.Cells[rowIndex, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[rowIndex, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                worksheet.Cells[rowIndex, headers.Length - 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            }
        }

        public static void ResetSubtotalVariables(ref int totalCount, ref double totalRecon, ref double totalTransFee, ref string totalCr, ref string totalDr)
        {
            totalCount = 0;
            totalRecon = 0;
            totalTransFee = 0;
            totalCr = "";
            totalDr = "";
        }

        public static void ResetGrandTotalVariables(ref int grandTotalCount, ref double grandTotalRecon, ref double grandTotalTransFee)
        {
            grandTotalCount = 0;
            grandTotalRecon = 0;
            grandTotalTransFee = 0;
        }

        public static void ResetSubtotalRejectVaribles(ref double sourceAmount)
        {
            sourceAmount = 0;
        }

        public static void ResetGrandtoalRejectVariables(ref double grandTotalSourceAmount)
        {
            grandTotalSourceAmount = 0;
        }
    }
}
