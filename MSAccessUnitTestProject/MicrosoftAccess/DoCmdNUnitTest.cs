using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
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

        [Test]
        public void GetProcessId_Test()
        {

            var microsoftAccess = new Microsoft.Office.Interop.Access.Application();
            microsoftAccess.OpenCurrentDatabase(@"D:\MSAccessDatabase\MSAccessDatabase.accdb");

            var myName = microsoftAccess.Run("GetName");

            int id;
            GetWindowThreadProcessId(microsoftAccess.hWndAccessApp(), out id);

            var result = Process.GetProcessById(id);

            Assert.IsNotNull(result);

            Console.WriteLine(result.Id);

            if (microsoftAccess != null)
            {
                microsoftAccess.Quit();
                Marshal.ReleaseComObject(microsoftAccess);
                microsoftAccess = null;
            }
        }

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

    }
}
