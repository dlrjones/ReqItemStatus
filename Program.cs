using System;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using LogDefault;

namespace ReqItemStatus
{
    class Program
    {
        #region class variables
        private static string corpName = "";  //HMC, UWMC
        private static string logPath = "";
        private static bool debug = false;
        private static bool cleanUp = false;
        private static LogManager lm = LogManager.GetInstance();
        private static NameValueCollection ConfigData = null;
        #endregion
        static void Main(string[] args)
        {
           // check the cleanUp param
            if (args.Length > 0)
            {
                cleanUp = true;
            }
            //cleanUp = true;                                      
            try
            {
                GetParameters();
                lm.Write("Program/Main:  " + "BEGIN");                
                LoadData();
                if (cleanUp)
                    lm.Write("Program/Main:  " + "END");
                else
                {
                    RemoveAttachments();
                    Process();
                    lm.Write("Program/Main:  " + "END");
                }
            }
            catch (Exception ex)
            {
                lm.Write("Program/Main:  " + ex.Message);
            }
            finally
            {
                Environment.Exit(1);
            }
        }

        private static void GetParameters()
        {
            ConfigData = (NameValueCollection)ConfigurationSettings.GetConfig("appSettings");
            debug = Convert.ToBoolean(ConfigData.Get("debug"));
            lm.LogFilePath = ConfigData.Get("logFilePath");
            lm.LogFile = ConfigData.Get("logFile");
            lm.Debug = Convert.ToBoolean(ConfigData.Get("debug"));
        }
     
        private static void LoadData()
        {
            DataSetManager dsm = DataSetManager.GetInstance();          
            dsm.Debug = debug;
            if (cleanUp)
            {//removes REQ_ITEM_ID's that have a status of KILLED, COMPLETE or REMOVED (Denied). This needs to be run once each day after midnight.
                //The initial select query (DataSetManager.BuildTodayQuery) only looks for records from the previous run through to the current run time.
                //To run DeleteKilledComplete, launch ReqItemStatus wih the number "1" as a parameter (or anything, really - as you can see above 
                //it only looks at the number of arguments over 0). You'll probably want this as its own Scheduled Task (currently set at 12:30AM).
                lm.Write("Program/Main.LoadData:  " + "CleanUp - hmcmm_ReqItemStatus");
                dsm.DeleteKilledComplete();
            }
            else
            {
                dsm.LoadTodaysDataSet();
                dsm.LoadYesterdayList();
            }
        }
        
        private static void RemoveAttachments()
        {
            //Excel files are created with the req item status data and then saved to be 
            //later attached to the email sent to each req creator. This method removes the 
            //attachment files from the previous run

            string attachmentPath = ConfigData.Get("attachmentPath");
            string[] files = Directory.GetFiles(attachmentPath, "*.xlsx");
            int fileCount = 0;
            try
            {
                foreach (string fName in files)
                {
                    File.Delete(fName);
                    fileCount++;
                }
                if(fileCount == 1)
                    lm.Write(fileCount + " old attachment deleted.");
                else
                    lm.Write(fileCount + " old attachments deleted.");
            }
            catch (Exception ex)
            {
                lm.Write("Program/RemoveAttachments:  " + ex.Message);
            }
        }

        private static void Process()
        {
            ProcessManager pm = new ProcessManager();            
            pm.Debug = debug;
            pm.Begin();
        }
    
    }
}
