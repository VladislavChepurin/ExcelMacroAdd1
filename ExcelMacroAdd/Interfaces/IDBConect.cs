using ExcelMacroAdd.UserVariables;

namespace ExcelMacroAdd.Interfaces
{
    interface IDBConect
    {        
        void OpenDB();
        void CloseDB();
        string ReadOnlyOneNoteDB(string requestDB, int colum);
        void UpdateNotesDB(string dataRequest);
        DBtable ReadSeveralNotesDB(string dataRead);
    }
}
