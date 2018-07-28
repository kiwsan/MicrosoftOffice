using log4net;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsService
{

    //http://docs.pdfforge.org/pdfcreator/3.2/en/pdfcreator/com-interface/
    //https://stackoverflow.com/questions/50100556/printing-excel-sheet-using-pdfcreator-net-com-wrapper
    public static class PDFCreatorHelper
    {
        static readonly ILog Logger = LogManager.GetLogger("Program");

        //public static string ErrorText { get; private set; }
        //public static string CreatedFile { get; private set; }
        //private static pdfforge.PDFCreator.UI.ComWrapper.Queue jobQueue = null;

        //public static void PrintSheet(Worksheet xlSheet, Application app, string file)
        //{

        //    ErrorText = null;
        //    CreatedFile = null;
        //    if (jobQueue == null)
        //    {
        //        Type queueType = Type.GetTypeFromProgID("PDFCreator.JobQueue");
        //        Activator.CreateInstance(queueType);
        //        jobQueue = new pdfforge.PDFCreator.UI.ComWrapper.Queue();
        //        jobQueue.Initialize();  //Reusing one instance for the application runtime
        //    }
        //    else
        //    {
        //        jobQueue.Clear(); //Delete jobs already put there
        //    }
        //    string folder = Path.GetDirectoryName(file);
        //    string filename = Path.GetFileName(file);
        //    string convertedFilePath = Path.Combine(folder, filename);

        //    try
        //    {
        //        //Actual print command
        //        xlSheet.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, Type.Missing, "PDFCreator", Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        //        if (!jobQueue.WaitForJob(10))
        //        {
        //            ErrorText = string.Format("PDFCreator: tisk {0} nebyl spuštěn do 10 sekund.", file);
        //            Logger.Error(ErrorText);
        //        }
        //        else
        //        {
        //            var printJob = jobQueue.NextJob;
        //            printJob.SetProfileByGuid("DefaultGuid");
        //            printJob.SetProfileSetting("OpenViewer", "false");
        //            printJob.SetProfileSetting("OpenWithPdfArchitect", "false");
        //            printJob.SetProfileSetting("ShowProgress", "false");
        //            printJob.SetProfileSetting("TargetDirectory", folder);
        //            printJob.SetProfileSetting("ShowAllNotifications", "false");

        //            if (File.Exists(convertedFilePath)) File.Delete(convertedFilePath);
        //            printJob.ConvertTo(convertedFilePath);

        //            if (!printJob.IsFinished || !printJob.IsSuccessful)
        //            {
        //                ErrorText = string.Format("PDFCreator: nepodařila se konverze souboru: {0}.", file);
        //                Logger.Error(ErrorText);
        //            }
        //            printJob = null;
        //        }
        //    }
        //    catch (Exception err)
        //    {
        //        Logger.Error(err);
        //        ErrorText = err.Message;
        //    }
        //    finally
        //    {
        //        CreatedFile = convertedFilePath;
        //        //If this is left uncommented, app hangs during the second run
        //        //jobQueue.ReleaseCom(); 
        //        //jobQueue = null;
        //    }
        //}

    }
}
