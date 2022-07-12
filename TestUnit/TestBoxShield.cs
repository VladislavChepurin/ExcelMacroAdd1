using BenchmarkDotNet.Attributes;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Services;
using ExcelMacroAdd.Servises;
using System;

namespace TestUnit
{
    internal class TestBoxShield
    {
        [Benchmark]
        public void StartTest()
        {
            Lazy<DataInXml> dataInXml = new Lazy<DataInXml>();
            DBConectProxy dBConect = new DBConectProxy(new Lazy<DBConect>());

            BoxShield boxShield = new BoxShield(dBConect);
            boxShield.Start();
        }

    }
}
