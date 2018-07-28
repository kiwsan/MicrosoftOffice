using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MSAccessUnitTestProject.MicrosoftAccess
{
    [TestFixture]
    public class DoCmdNUnitTest
    {

        [Test]
        public void GetName()
        {

            var microsoftAccess = new Microsoft.Office.Interop.Access.Application();
            microsoftAccess.Visible = true;
            microsoftAccess.OpenCurrentDatabase(@"D:\MSAccessDatabase.accdb", false);

            var myName = microsoftAccess.Run("GetName");

            Console.WriteLine($"My Name: {myName}");

            Thread.Sleep(3000);

            microsoftAccess.CloseCurrentDatabase();
            microsoftAccess.Quit();

        }

    }
}
