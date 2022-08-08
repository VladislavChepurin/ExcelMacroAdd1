﻿using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.Servises;
using ExcelMacroAdd.UserVariables;
using Moq;
using NUnit.Framework;
using System;
using System.IO;
using System.Security;

namespace ExelMacroAdd.Tests
{
    [TestFixture]
    public class DataInXmlTest
    {
        [Test]
        public void ReadElementXmlMustNotNull()
        {
            var data = new DataInXml();            
            Assert.IsNotNull(data.ReadElementXml("IEK", data.ReadFileXml()));
            Assert.IsNotNull(data.ReadElementXml("EKF", data.ReadFileXml()));
            Assert.IsNotNull(data.ReadElementXml("DKC", data.ReadFileXml()));
            Assert.IsNotNull(data.ReadElementXml("KEAZ", data.ReadFileXml()));
            Assert.IsNotNull(data.ReadElementXml("DEKraft", data.ReadFileXml()));
            Assert.IsNotNull(data.ReadElementXml("TDM", data.ReadFileXml()));
            Assert.IsNotNull(data.ReadElementXml("ABB", data.ReadFileXml()));
            Assert.IsNotNull(data.ReadElementXml("Schneider", data.ReadFileXml()));
        }
        [Test]
        public void ReadElementXmlMustCorrectValue()
        {
            var testVendorObject = new Vendor[] {            
                new Vendor
                {
                    VendorAttribute = "IEK",
                    Date = "2020-05-04 15:14:45",
                    Discont = 25,
                    Formula_1 = @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$K$65536;2;0)",
                    Formula_2 = @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$K$65536;4;0)",
                    Formula_3 = @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$J$65536;10;0)"
                },
                new Vendor
                {
                    VendorAttribute = "DKC",
                    Date = "2019-01-12 12:34:11",
                    Discont = 45,
                    Formula_1 = @"=ВПР(A11;'C:\Users\ПК\Desktop\Прайсы\[DKC-Prays_list-ot-14.01.2022.xlsx]Прайс ДКС'!$F$15:$M$65536;2;0)",
                    Formula_2 = @"=ВПР(A11;'C:\Users\ПК\Desktop\Прайсы\[DKC-Prays_list-ot-14.01.2022.xlsx]Прайс ДКС'!$F$15:$M$65536;3;0)",
                    Formula_3 = @"=ВПР(A11;'C:\Users\ПК\Desktop\Прайсы\[DKC-Prays_list-ot-14.01.2022.xlsx]Прайс ДКС'!$F$15:$N$65536;5;0)"
                }
            };
            
            var mock = new Mock<IDataInXml>();
            mock.Setup(p => p.ReadFileXml()).Returns(testVendorObject);
            
            var data = new DataInXml();
            Assert.AreEqual(data.ReadElementXml("IEK", mock.Object.ReadFileXml()).VendorAttribute, "IEK");
            Assert.AreEqual(data.ReadElementXml("IEK", mock.Object.ReadFileXml()).Discont, 25);
            Assert.AreEqual(data.ReadElementXml("IEK", mock.Object.ReadFileXml()).Date, "2020-05-04 15:14:45");
            Assert.AreEqual(data.ReadElementXml("IEK", mock.Object.ReadFileXml()).Formula_1, @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$K$65536;2;0)");
            Assert.AreEqual(data.ReadElementXml("IEK", mock.Object.ReadFileXml()).Formula_2, @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$K$65536;4;0)");
            Assert.AreEqual(data.ReadElementXml("IEK", mock.Object.ReadFileXml()).Formula_3, @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$J$65536;10;0)");

            Assert.AreEqual(data.ReadElementXml("DKC", mock.Object.ReadFileXml()).VendorAttribute, "DKC");
            Assert.AreEqual(data.ReadElementXml("DKC", mock.Object.ReadFileXml()).Discont, 45);
            Assert.AreEqual(data.ReadElementXml("DKC", mock.Object.ReadFileXml()).Date, "2019-01-12 12:34:11");
            Assert.AreEqual(data.ReadElementXml("DKC", mock.Object.ReadFileXml()).Formula_1, @"=ВПР(A11;'C:\Users\ПК\Desktop\Прайсы\[DKC-Prays_list-ot-14.01.2022.xlsx]Прайс ДКС'!$F$15:$M$65536;2;0)");
            Assert.AreEqual(data.ReadElementXml("DKC", mock.Object.ReadFileXml()).Formula_2, @"=ВПР(A11;'C:\Users\ПК\Desktop\Прайсы\[DKC-Prays_list-ot-14.01.2022.xlsx]Прайс ДКС'!$F$15:$M$65536;3;0)");
            Assert.AreEqual(data.ReadElementXml("DKC", mock.Object.ReadFileXml()).Formula_3, @"=ВПР(A11;'C:\Users\ПК\Desktop\Прайсы\[DKC-Prays_list-ot-14.01.2022.xlsx]Прайс ДКС'!$F$15:$N$65536;5;0)");
        }
        [Test]
        public void ReadElementNotMustTrowsException()
        {
            var testVendorObject = new Vendor[] {
                new Vendor
                {
                    VendorAttribute = "IEK",
                    Date = "2020-05-04 15:14:45",
                    Discont = 25,
                    Formula_1 = @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$K$65536;2;0)",
                    Formula_2 = @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$K$65536;4;0)",
                    Formula_3 = @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$J$65536;10;0)"
                },
                new Vendor
                {
                    VendorAttribute = "IEK",
                    Date = "2020-05-04 15:14:45",
                    Discont = 25,
                    Formula_1 = @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$K$65536;2;0)",
                    Formula_2 = @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$K$65536;4;0)",
                    Formula_3 = @"=ВПР(A3;'C:\Users\ПК\Desktop\Прайсы\[220131 Прайс.xlsx]Прайс'!$A$13:$J$65536;10;0)"
                }
            };

            var mock = new Mock<IDataInXml>();
            mock.Setup(p => p.ReadFileXml()).Returns(testVendorObject);

            var data = new DataInXml();            
            Assert.Throws<InvalidOperationException>(() => data.ReadElementXml("IEK", mock.Object.ReadFileXml()));          
        }          
    }
}
