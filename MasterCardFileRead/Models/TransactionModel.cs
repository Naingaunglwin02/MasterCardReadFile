namespace MasterCardFileRead.Models
{
    public class TransactionModel
    {
        public string Date { get; set; }

        public string? MemberID { get; set; }

        public string? FileId { get; set; }

        public string? TranscFunction { get; set; }

        public List<string>? Ird { get; set; }

        public List<string>? Count { get; set; }

        public List<string>? ReconAmount { get; set; }

        public List<string>? TransferFee { get; set; }

        public List<string>? Code { get; set; }

        public List<string>? Proc { get; set; }

        public string? EndOfReport { get; set; }

        public string? Cycle { get; set; }

    }
}
