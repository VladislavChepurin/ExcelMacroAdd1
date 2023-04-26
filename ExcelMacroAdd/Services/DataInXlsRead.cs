using ExcelMacroAdd.UserVariables;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Markup;

namespace ExcelMacroAdd.Services
{
    public class DataInXlsRead
    {
        private readonly string BaseDirectory;

        public DataInXlsRead()
        {
            BaseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config");
        }
 

        public Vendor ReadElementXls(string vendor, Vendor[] dataXmlContinue)
        {
            return dataXmlContinue.Single(p => p.VendorAttribute == ReplaceVendorTable()[vendor]);         
        }

        public Vendor[] GetReadVendors()
        {
            throw new NotImplementedException();
        }

        public string GetFileXlsNameInDirectory(string dirName)
        {
            if (Directory.Exists(dirName))
            {
                try
                {
                    var files = Directory.GetFiles(dirName)
                        .Where(f => f.Contains(".xlsx") || f.Contains(".xls")).AsEnumerable();
                    var maxDateTime = files.Max(data => File.GetCreationTime(data));
                    var filePath = files.FirstOrDefault(data => File.GetCreationTime(data) == maxDateTime);
                    return filePath;
                }
                catch (InvalidOperationException)
                {
                    return null;
                }    
            }
            return null;
        }

        /// <summary>
        /// Функция замены для вставки вендора и запроса из XML
        /// </summary>
        /// <returns></returns>
        private IDictionary<string, string> ReplaceVendorTable()
        {
            Dictionary<string, string> dictionaryVendor = new Dictionary<string, string>()
            {
                {"Iek", "IEK"},
                {"Ekf", "EKF"},
                {"IekVa47", "IEK"},
                {"IekVa47m", "IEK"},
                {"IekArmat", "IEK"},
                {"EkfProxima", "EKF"},
                {"EkfAvers", "EKF"},
                {"Abb", "ABB"},
                {"Keaz", "KEAZ"},
                {"Dkc", "DKC"},
                {"Dekraft", "DEKraft"},
                {"Schneider", "Schneider"},
                {"Chint", "Chint"},
                {"Tdm", "TDM"}
            };
            return dictionaryVendor;
        }
    }
}
