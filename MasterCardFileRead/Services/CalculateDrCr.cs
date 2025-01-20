using MasterCardFileRead.Models;
using OfficeOpenXml;

namespace MasterCardFileRead.Services
{
    public class CalculateDrCr
    {
        public static bool ShouldAddSubtotals(string previousCycle, string currentCycle)
        {
            return previousCycle != null && currentCycle != previousCycle;
        }

        public static bool ShouldAddGrandTotals(string previousDate, string currentDate)
        {
            return previousDate != null && currentDate != previousDate;
        }

        public static void AddSubtotals(ExcelWorksheet worksheet, ref int rowIndex, ref int totalCount, ref double totalDr, ref double totalCr, ref double totalTranDr, ref double totalTranCr, ref double grandTotalDr, ref double grandTotalCr, ref double grandTotalTransDr, ref double grandTotalTransCr)
        {
            var subTotal = Math.Abs(totalDr - totalCr);
            var subTotalType = totalDr > totalCr ? "DR" : "CR";

            UpdateGrandTotals(subTotal, subTotalType, ref grandTotalDr, ref grandTotalCr);

            var subTotalTransFee = Math.Abs(totalTranDr - totalTranCr);
            var subTotalTransFeeType = totalTranDr > totalTranCr ? "DR" : "CR";

            UpdateGrandTotals(subTotalTransFee, subTotalTransFeeType, ref grandTotalTransDr, ref grandTotalTransCr);

            TotalTransactions.AddSubtotalRow(worksheet, rowIndex, "Total", totalCount, subTotal, subTotalTransFee, subTotalType, subTotalTransFeeType);
            rowIndex += 2;

            TotalTransactions.ResetSubtotalVariables(ref totalCount, ref totalCr, ref totalDr, ref totalTranCr, ref totalTranDr);
        }

        public static void AddGrandTotals(ExcelWorksheet worksheet, ref int rowIndex, int grandTotalCount, double grandTotalDr, double grandTotalCr, double grandTotalTransDr, double grandTotalTransCr)
        {
            double dateGrandTotal = Math.Abs(grandTotalDr - grandTotalCr);
            string dateGrandTotalType = grandTotalDr > grandTotalCr ? "DR" : "CR";

            double dateGrandTransTotal = Math.Abs(grandTotalTransDr - grandTotalTransCr);
            string dateGrandTransTotalType = grandTotalTransDr > grandTotalTransCr ? "DR" : "CR";

            TotalTransactions.AddGrandTotalRow(
                worksheet,
                rowIndex,
                "Grand Total",
                grandTotalCount,
                dateGrandTotal,
                dateGrandTransTotal,
                dateGrandTotalType,
                dateGrandTransTotalType
            );
            rowIndex += 2;
        }

        public static void WriteTransactionData(ExcelWorksheet worksheet, TransactionModel record, int rowIndex, ref double totalDr, ref double totalCr, ref double totalTranDr, ref double totalTranCr, ref int totalCount, ref int grandTotalCount)
        {
            var partReconAmount = record.ReconAmount.Split(" ");
            var partTransferFee = record.TransferFee.Split(" ");
            if (partReconAmount.Length > 0)
            {
                worksheet.Cells[rowIndex, 10].Value = partReconAmount[0];
            }
            if (partTransferFee.Length > 0)
            {
                worksheet.Cells[rowIndex, 13].Value = partTransferFee[0];
            }

            worksheet.Cells[rowIndex, 1].Value = record.TranscFunction;
            worksheet.Cells[rowIndex, 2].Value = record.Date;
            worksheet.Cells[rowIndex, 3].Value = record.FileId;
            worksheet.Cells[rowIndex, 4].Value = record.MemberID;
            worksheet.Cells[rowIndex, 5].Value = record.Cycle;
            worksheet.Cells[rowIndex, 6].Value = record.Proc;
            worksheet.Cells[rowIndex, 7].Value = record.Code;
            worksheet.Cells[rowIndex, 8].Value = record.Ird;
            worksheet.Cells[rowIndex, 9].Value = record.Count;
            worksheet.Cells[rowIndex, 11].Value = record.ReconDCCR;
            worksheet.Cells[rowIndex, 12].Value = record.Currency;
            worksheet.Cells[rowIndex, 14].Value = record.TransferFeeDCCR;

            worksheet.Cells.AutoFitColumns();

            ParseAndAccumulateValues(record.ReconAmount, ref totalDr, ref totalCr);
            ParseAndAccumulateValues(record.TransferFee, ref totalTranDr, ref totalTranCr);

            totalCount += int.Parse(record.Count);
            grandTotalCount += int.Parse(record.Count);
        }

        private static void ParseAndAccumulateValues(string value, ref double drValue, ref double crValue)
        {
            var parts = value.Split(" ");
            if (parts.Length == 2)
            {
                double amount = Convert.ToDouble(parts[0]);
                if (parts[1] == "DR")
                {
                    drValue += amount;
                }
                else if (parts[1] == "CR")
                {
                    crValue += amount;
                }
            }
        }

        public static void AddFinalSubtotalsAndGrandTotals(ExcelWorksheet worksheet, ref int rowIndex, string previousCycle, int totalCount, double totalDr, double totalCr, double totalTranDr, double totalTranCr, ref double grandTotalDr, ref double grandTotalCr, ref double grandTotalTransDr, ref double grandTotalTransCr, int grandTotalCount)
        {
            if (previousCycle != null)
            {
                AddSubtotals(worksheet, ref rowIndex, ref totalCount, ref totalDr, ref totalCr, ref totalTranDr, ref totalTranCr, ref grandTotalDr, ref grandTotalCr, ref grandTotalTransDr, ref grandTotalTransCr);
            }

            if (grandTotalCount > 0)
            {
                AddGrandTotals(worksheet, ref rowIndex, grandTotalCount, grandTotalDr, grandTotalCr, grandTotalTransDr, grandTotalTransCr);
            }
        }

        private static void UpdateGrandTotals(double amount, string type, ref double drValue, ref double crValue)
        {
            if (type == "DR")
            {
                drValue += amount;
            }
            else
            {
                crValue += amount;
            }
        }
    }
}
