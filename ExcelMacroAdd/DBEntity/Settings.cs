
using System.Data.Linq.Mapping;


namespace ExcelMacroAdd.DBEntity
{
    [Table(Name ="settings")]
    public class Settings
    {
        [Column(IsPrimaryKey =true, IsDbGenerated =true)]
        public int Id { get; set; }

        [Column(Name ="set_name")]
        public string SetName { get; set; }

        [Column(Name ="set_option")]
        public string SetOption { get; set; }

    }
}
