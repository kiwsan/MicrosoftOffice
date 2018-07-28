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
    partial class WordPrintService : ServiceBase
    {

        // Two Seconds  
        private Timer timerTwoSeconds = new Timer(2000);
        private WordPrintTask wordPrintTask = new WordPrintTask();
        public WordPrintService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // TODO: Add code here to start your service.
            timerTwoSeconds.Elapsed += new ElapsedEventHandler(TimerTwoSeconds_Elapsed);
            timerTwoSeconds.Enabled = true;
        }

        private void TimerTwoSeconds_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                Task task1 = Task.Factory.StartNew(() => wordPrintTask.PrintWord());
            }
            catch (Exception ex)
            {
            }
        }

        protected override void OnStop()
        {
            // TODO: Add code here to perform any tear-down necessary to stop your service.
            timerTwoSeconds.Stop();
        }

    }
}
