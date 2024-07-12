using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelMacroAdd.Functions
{
    internal sealed class YandexSearch : AbstractFunctions
    {
        public override void Start()
        {
            var value = Cell.Value;
            if (value != null)
            {
                string request;
                if (value is Object[,])
                {
                    List<string> list = new List<string>();
                    foreach (var item in value)
                    {
                        if (item != null)
                            list.Add(item.ToString());
                    }
                    request = String.Join(" ", list);
                }
                else
                {
                    request = Cell.Value.ToString();
                }                    
                string url = "http://www.yandex.ru/yandsearch?text=" + request; 
                Process.Start(url);
            }                             
        }
    }
}
