using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.Servises;
using ExcelMacroAdd.UserVariables;
using System.Collections.Generic;

namespace ExcelMacroAdd.Services
{
    internal class DBConectProxy : IDBConect
    {
        private readonly IDictionary<string, string> _cacheOnlyOneNoteDB = new Dictionary<string, string>();
        private readonly IDictionary<string, DBtable> _cacheSeveralNotesDB = new Dictionary<string, DBtable>();
        private readonly DBConect _dBConect;  
        
        public DBConectProxy(DBConect dBConect)
        {
            _dBConect = dBConect;
        }
               
        public string PPatch => _dBConect.PPatch;
        public string SPatch => _dBConect.SPatch;
        public string ProviderData => _dBConect.ProviderData;

        public void OpenDB()
        {
            //Проксируем подключение
            _dBConect.OpenDB();
        }

        public void CloseDB()
        {
            //Проксируем отключение
            _dBConect.CloseDB();
        }

        public string ReadOnlyOneNoteDB(string requestDB, int colum)
        {
            if (!_cacheOnlyOneNoteDB.ContainsKey(requestDB))
            {                
                var value = _dBConect.ReadOnlyOneNoteDB(requestDB, colum);
                _cacheOnlyOneNoteDB.Add(requestDB, value);
                return value;
            }
            return _cacheOnlyOneNoteDB[requestDB];
        }

        public void UpdateNotesDB(string queryUpdate, string dataRequest)
        {
            _cacheSeveralNotesDB.Clear();
            //обновление записей проксируем напрямую
            _dBConect.UpdateNotesDB(queryUpdate, dataRequest);
        }
      
        public DBtable ReadSeveralNotesDB(string dataRead)
        {
            if (!_cacheSeveralNotesDB.ContainsKey(dataRead))
            {
                var value = _dBConect.ReadSeveralNotesDB(dataRead);
                _cacheSeveralNotesDB.Add(dataRead, value);
                return value;
            }
            return _cacheSeveralNotesDB[dataRead];
        }
    }
}
 