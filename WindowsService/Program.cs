using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WindowsService
{
    static class Program
    {

        static readonly ILog Logger = LogManager.GetLogger("Program");
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {

            //#if DEBUG
            //            Service1 myService = new Service1();
            //            myService.OnDebug();
            //            System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);
            //#endif
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new Service1()
            };
            ServiceBase.Run(ServicesToRun);

        }

    }
}
