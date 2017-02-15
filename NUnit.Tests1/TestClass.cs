using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenOfficeWpfApp;
using static DocumentFormat.OpenXml.Packaging.WordprocessingDocument;

namespace NUnit.Tests1
{
    [TestFixture]
    public class TestClass
    {
        [Test]
        public void TestChangePackageMethod()
        {
            var testOfficePackage = new ChangeTestOfficeFileClass();
            testOfficePackage.ChangePackage(@"C:\Backup\test6.docx");
            //Assert.Pass("ChangePackage passing test (no exception");
        }

        //[Test]
        //public void TestAddPersonalInfoMethod()
        //{
        //    using (var document = Open(@"C:\Backup\test4.docx", true))
        //    {
        //        Document document1 = document.MainDocumentPart.Document;
        //        Body body = document1.GetFirstChild<Body>();

        //        AddPersonalInfo(body);
        //    }
        //}
    }
}