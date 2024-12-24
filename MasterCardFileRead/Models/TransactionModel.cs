namespace MasterCardFileRead.Models
{
    public class TransactionModel
    {
        public string Date { get; set; }

        public string? MemberID { get; set; }

        public string? FileId { get; set; }

        public string? TranscFunction { get; set; }

        public string? Ird { get; set; }

        public string? Count { get; set; }

        public string? ReconAmount { get; set; }

        public string? TransferFee { get; set; }

        public string? Code { get; set; }

        public string? Proc { get; set; }

        public string? EndOfReport { get; set; }

        public string? Cycle { get; set; }

        public string? Currency { get; set; }

        public string? ReconDCCR { get; set; }

        public string? TransferFeeDCCR { get; set; }

    }

    public class TransactionResult
    {
        public string? TransactionFunction { get; set; }
        public string? Code { get; set; }
        public string? IrdValues { get; set; }
        public string? Count { get; set; }
        public string? ReconAmount { get; set; }
        public string? TransferFee { get; set; }
        public string? Proc { get; set; }
        public string? Currency { get; set; }
        public string? ReconDCCR { get; set; }
        public string? TransferFeeDCCR { get; set; }
    }

    public class RejectTransactionModel
    {
        //public string? Date { get; set; }
        public string? ProcessingMode { get; set; }
        public string? MtiFunctionCode { get; set; }
        public string? FileId { get; set; }
        public List<string>? ErrorCode { get; set; }
        public List<string>? ErrorDescription { get; set; }
        public string? SourceMessage { get; set; }
        public List<string>? ElementId { get; set; }
        public string? CardNumberD0002 { get; set; }
        public string? MccCodeD0026 { get; set; }
        public string? RrnD0037 { get; set; }
        public string? AuthCodeD0038 { get; set; }
        //public string? TerminalId { get; set; }
        //public string? MerchantIdD0042 { get; set; }
        //public string? MerchantNameD0043S01 { get; set; }
        //public string? IrdP0158S04 { get; set; }
        //public string? SourceAmount { get; set; }
        //public string? SourceCurrency { get; set; }

    }
}
