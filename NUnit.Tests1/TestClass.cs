using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenOfficeWpfApp;

namespace NUnit.Tests1
{
    [TestFixture]
    public class TestClass
    {
        [Test]
        public void TestChangePackageMethod()
        {
            var testOfficePackage = new ChangeTestOfficeFileClass();
            testOfficePackage.ChangePackage(@"C:\Backup\test4.docx");
            //Assert.Pass("ChangePackage passing test");
        }
    }
}
