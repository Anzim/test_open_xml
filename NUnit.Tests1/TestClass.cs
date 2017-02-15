using NUnit.Framework;
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
            testOfficePackage.ChangePackage(@"C:\Backup\test6.docx");
            //Assert.Pass("ChangePackage passing test (no exception");
        }

    }
}