using ExcelMacroAdd.Services;
using NUnit.Framework;

namespace ExelMacroAdd.Tests
{
    [TestFixture]
    public class DataInXlsReadTest
    {
        [Test]
        public void ReadFileExcelMustNotNull()
        {
            var data = new DataInXlsRead();
            Assert.IsNotNull(data.GetFileXlsNameInDirectory(@"D:\\Test"));          
        }

        [Test]
        public void ReadFileExcelMustNotException()
        {
            var data = new DataInXlsRead(); 
            Assert.DoesNotThrow(() => { data.GetFileXlsNameInDirectory(@"D:\\Test"); });
        }
    }
}
