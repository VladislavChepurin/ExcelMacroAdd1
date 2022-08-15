using ExcelMacroAdd.UserVariables;

namespace ExcelMacroAdd.Interfaces
{
    interface IDBConect
    {
        string PPatch { get; }
        string SPatch { get; }
        string ProviderData { get; }
        void OpenDB();
        void CloseDB();
        string ReadOnlyOneNoteDB(string requestDB, int colum);
        void UpdateNotesDB(string queryUpdate, string dataRequest);
        DBtable ReadSeveralNotesDB(string dataRead);
    }
}
