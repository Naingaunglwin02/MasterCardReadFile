namespace MasterCardFileRead.Models
{
    public class PresModel
    {
        public string? RecordType { get; set; }

        public string DistributionID { get; set; }

        public string Date { get; set; }

        public string? MemberID { get; set; }

        public string PageNo { get; set; }

        public string? AcceptanceBrand { get; set; }

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
