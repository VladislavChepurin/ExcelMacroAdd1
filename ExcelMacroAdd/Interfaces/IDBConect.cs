using ExcelMacroAdd.UserVariables;
using System;

namespace ExcelMacroAdd.Interfaces
{
    public interface IDBConect
    {
        string PPatch { get;}
        void OpenDB();
        void CloseDB();
        string ReadOnlyOneNoteDB(string requestDB, int colum);
        void UpdateNotesDB(string queryUpdate, string dataRequest);
        DBtable ReadSeveralNotesDB(string dataRead);

    }
}
