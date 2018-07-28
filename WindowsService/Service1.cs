using log4net;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace WindowsService
{

    public partial class Service1 : ServiceBase
    {

        static readonly ILog Logger = LogManager.GetLogger("Service1");

        // Two Seconds  
        private Timer timerTwoSeconds = new Timer(20000);
        private AccessPrintTask accessPrintTask = new AccessPrintTask();

        internal void OnDebug()
        {
            OnStart(null);
        }

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // TODO: Add code here to start your service.
            Logger.Info("OnStart..");
            timerTwoSeconds.Elapsed += new ElapsedEventHandler(TimerTwoSeconds_Elapsed);
            timerTwoSeconds.Enabled = true;
        }

        private void TimerTwoSeconds_Elapsed(object sender, ElapsedEventArgs e)
        {
            Logger.Info("Timestamp..");
            try
            {
                Task task1 = Task.Factory.StartNew(() => accessPrintTask.Start());
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
        }

        protected override void OnStop()
        {
            // TODO: Add code here to perform any tear-down necessary to stop your service.
            Logger.Info("OnStop..");
            timerTwoSeconds.Stop();
        }

    }
}
