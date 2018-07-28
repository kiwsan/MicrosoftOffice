using log4net;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WindowsService
{
    partial class AccessPrintService : ServiceBase
    {
        static readonly ILog Logger = LogManager.GetLogger("Program");

        private Microsoft.Office.Interop.Access.Application microsoftAccess = null;
        private System.Threading.Timer workTimer;    // System.Threading.Timer
        internal void OnDebug()
        {
            OnStart(null);
        }

        public AccessPrintService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Logger.Info("OnStart..");

            workTimer = new System.Threading.Timer(new TimerCallback(DoWork), null, 5000, 30000);

            base.OnStart(args);
        }

        private void DoWork(object state)
        {
            Logger.Info("Timestamp..");


        }

        protected override void OnPause()
        {
            workTimer.Change(Timeout.Infinite, Timeout.Infinite);
            base.OnPause();
        }

        protected override void OnContinue()
        {
            workTimer.Change(0, 30000);
            base.OnContinue();
        }

        protected override void OnStop()
        {
            Logger.Info("OnStop..");
            workTimer.Dispose();
            base.OnStop();


            //if (microsoftAccess != null)
            //{
            //    microsoftAccess.CloseCurrentDatabase();
            //    microsoftAccess.Quit();
            //}
        }
    }
}
