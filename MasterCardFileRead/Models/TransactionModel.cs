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

        //public List<string>? Proc { get; set; }

        public string? EndOfReport { get; set; }

        public string? Cycle { get; set; }

        //public string? Currency { get; set; }

    }

    public class TransactionResult
    {
        public string? TransactionFunction { get; set; }
        public string? Code { get; set; }
        public string? IrdValues { get; set; }
        public string? Count { get; set; }
        public string? ReconAmount { get; set; }
        public string? TransferFee { get; set; }
    }
}
