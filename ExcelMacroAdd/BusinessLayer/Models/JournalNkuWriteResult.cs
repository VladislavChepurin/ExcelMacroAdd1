namespace ExcelMacroAdd.BusinessLayer.Models
{
    public enum JournalNkuWriteStatus
    {
        Added,
        Updated,
        AlreadyExists,
        NotFound
    }

    public sealed class JournalNkuWriteResult
    {
        public JournalNkuWriteResult(JournalNkuWriteStatus status, string article)
        {
            Status = status;
            Article = article;
        }

        public JournalNkuWriteStatus Status { get; }

        public string Article { get; }
    }
}
