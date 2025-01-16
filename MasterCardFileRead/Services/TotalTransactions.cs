using OfficeOpenXml;
using System.Reflection.PortableExecutable;

namespace MasterCardFileRead.Services
{
    public class TotalTransactions
    {
        public static void AddSubtotalRow(ExcelWorksheet worksheet, int rowIndex, string title, int totalCount, double subTotal, double subTotalTransfee, string subTotalType, string subTotalTransFeeType)
        {
            worksheet.Cells[rowIndex, 1, rowIndex, 8].Merge = true;
            worksheet.Cells[rowIndex, 1].Value = title;

            worksheet.Cells[rowIndex, 9].Value = totalCount;
            worksheet.Cells[rowIndex, 10].Value = subTotal;
            worksheet.Cells[rowIndex, 10].Style.Numberformat.Format = "###0.00";
            worksheet.Cells[rowIndex, 11].Value = subTotalType;
            worksheet.Cells[rowIndex, 13].Value = subTotalTransfee;
            worksheet.Cells[rowIndex, 13].Style.Numberformat.Format = "###0.00";
            worksheet.Cells[rowIndex, 14].Value = subTotalTransFeeType;

            ApplySubTotalRowStyle(worksheet, rowIndex);
        }

        public static void AddGrandTotalRow(ExcelWorksheet worksheet, int rowIndex, string title, int grandTotalCount, double grandTotal, double grandTotalTransFee, string grandTotalType, string grandTotalTransFeeType)
        {
            worksheet.Cells[rowIndex, 1, rowIndex, 8].Merge = true;
            worksheet.Cells[rowIndex, 1].Value = title;

            worksheet.Cells[rowIndex, 9].Value = grandTotalCount;
            worksheet.Cells[rowIndex, 10].Value = grandTotal;
            worksheet.Cells[rowIndex, 10].Style.Numberformat.Format = "###0.00";
            worksheet.Cells[rowIndex, 11].Value = grandTotalType;
            worksheet.Cells[rowIndex, 13].Value = grandTotalTransFee;
            worksheet.Cells[rowIndex, 13].Style.Numberformat.Format = "###0.00";
            worksheet.Cells[rowIndex, 14].Value = grandTotalTransFeeType;

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

        public static void ResetSubtotalVariables(ref int totalCount, ref double totalCr, ref double totalDr, ref double totalTranCr, ref double totalTranDr)
        {
            totalCount = 0;
            totalTranCr = 0;
            totalTranDr = 0;
            totalCr = 0;
            totalDr = 0;
        }

        public static void ResetGrandTotalVariables(ref int grandTotalCount, ref double grandTotalCr, ref double grandTotalDr, ref double grandTotalTransDr, ref double grandTotalTransCr)
        {
            grandTotalCount = 0;
            grandTotalTransDr = 0;
            grandTotalTransCr = 0;
            grandTotalCr = 0;
            grandTotalDr = 0;
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
