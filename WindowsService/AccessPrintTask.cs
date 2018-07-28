using log4net;
using Microsoft.Office.Interop.Access;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsService
{

    //http://www.thaicreate.com/asp/forum/000303.html Error Cannot update Database or Object is read-only
    /*
     System.Runtime.InteropServices.COMException (0x800A9D9F): Exception from HRESULT: 0x800A9D9F at Microsoft.Office.Interop.Access.ApplicationClass.Run(String Procedure, Object& Arg1, Object &Arg2, ..., Object &Arg30)
         */
    //https://stackoverflow.com/questions/837754/call-routine-in-access-module-from-net
    //http://www.itgo.me/a/x438014096864530466/call-routine-in-access-module-from-net
    public class AccessPrintTask
    {
        static readonly ILog Logger = LogManager.GetLogger("Service1");
        private static object _locker = new Object();
        public AccessPrintTask()
        {

        }

        public void Start()
        {
            try
            {
                // Kill opened word instances.  
                if (KillProcess("MSACCESS"))
                {
                    // Thread safe.  
                    lock (_locker)
                    {
                        Logger.Info("Thread safe..");
                        string fileName = @"D:\MSAccessDatabase\MSAccessDatabase.accdb";
                        //string printerName = "PDFCreator";

                        if (File.Exists(fileName))
                        {
                            Logger.Info("Open..");
                            Application microsoftAccess = new Application();
                            microsoftAccess.OpenCurrentDatabase(fileName);

                            var myName = microsoftAccess.Run("GetName");

                            Logger.Info($"My Name: {myName}");

                            if (microsoftAccess != null)
                            {
                                microsoftAccess.Quit();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(microsoftAccess);
                                microsoftAccess = null;

                                Logger.Info("Quit..");
                            }
                        }
                    }

                }
            }
            catch (Exception ex )
            {
                Logger.Error(ex);
                Logger.Info("KillProcess..");
                KillProcess("MSACCESS");
            }
        }

        private static bool KillProcess(string name)
        {
            foreach (Process clsProcess in Process.GetProcesses().Where(p => p.ProcessName.Contains(name)))
            {
                if (Process.GetCurrentProcess().Id == clsProcess.Id)
                    continue;
                if (clsProcess.ProcessName.Contains(name))
                {
                    clsProcess.Kill();
                    return true;
                }
            }
            return true;
        }

    }
}
