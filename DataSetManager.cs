using System;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Threading;
using OleDBDataManager;
using LogDefault;

namespace ReqItemStatus
{
    class DataSetManager
    {
        #region class variables
        private DataSet dsCurrent;
        private DataSet dsYesterday;
        private DataSet dsCurrentChangeDates;
        private Hashtable itemsThatChanged = new Hashtable();
        private Hashtable itemReq = new Hashtable();
        private Hashtable reqItemPONO = new Hashtable();
       // private DataTable reqNo_PO = new DataTable();
        private static string connectStrHEMM = "";
        private static string connectStrBIAdmin = "";
        private static NameValueCollection ConfigData = null;
        protected static ODMDataFactory ODMDataSetFactory = null;
        private static DataSetManager dataMngr = null;
        private static LogManager lm = LogManager.GetInstance();
        private static bool debug = false;
        private ArrayList tableRecord = new ArrayList();
        #region parameters
        public DataSet DsYesterday
        {
            get { return dsYesterday; }
            set { dsYesterday = value; }
        }

        public DataSet DSCurrent
        {
            get { return dsCurrent; }
            set { dsCurrent = value; }
        }
        public Hashtable ReqItemPONO
        {
            get { return reqItemPONO; }
        }
        public DataSet DSCurrentChangeDates
        {
            get { return dsCurrentChangeDates; }
            set { dsCurrentChangeDates = value; }
        }
        public Hashtable ItemsThatChanged
        {
            set { itemsThatChanged = value; }
        }
        public bool Debug
        {
            set { debug = value; }
        }
        public Hashtable ItemReq
        {
            set { itemReq = value; }
        }       
        public ArrayList TableRecord
        {
            get { return tableRecord; }
            set { tableRecord = value; }
        }
        #endregion parameters
        #endregion

        public DataSetManager()
        {
            InitDataSetManager();
          //  InitReqNo_PO();
        }    

        private static void InitDataSetManager()
        {                        
            ODMDataSetFactory = new ODMDataFactory();
            ConfigData = (NameValueCollection)ConfigurationSettings.GetConfig("appSettings");            
            connectStrHEMM = ConfigData.Get("cnctHEMM_HMC"); 
            connectStrBIAdmin = ConfigData.Get("cnctBIAdmin_HMC");
        }
        //private void InitReqNo_PO()
        //{
        //    ReqNo_PO.Columns.Add("KEY", typeof(string));
        //    ReqNo_PO.Columns.Add("REQ_NO", typeof(string));
        //    ReqNo_PO.Columns.Add("PO_NO", typeof(string));
        //    ReqNo_PO.Columns.Add("ITEM_NO", typeof(string));
        //    ReqNo_PO.Columns.Add("PO_LINE", typeof(string));
        //}

        public static DataSetManager GetInstance()
        {
            if (dataMngr == null)
            {
                CreateInstance();
            }
            return dataMngr;
        }

        private static void CreateInstance()
        {
            Mutex configMutex = new Mutex();
            configMutex.WaitOne();
            dataMngr = new DataSetManager();
            configMutex.ReleaseMutex();
        }

        public void LoadTodaysDataSet()
        {
            ODMRequest Request = new ODMRequest();
            Request.ConnectString = connectStrBIAdmin;   //connectStrHEMM;
            Request.CommandType = CommandType.Text;
            Request.Command = "Execute ('" + BuildTodayQuery() + "')";

            if (debug)
                lm.Write("DataSetManager/LoadTodaysDataSet:  " + "BuildTodayQuery() - list of Req Items");   //Request.Command);
            try
            {
                dsCurrent = ODMDataSetFactory.ExecuteDataSetBuild(ref Request);
            }
            catch (Exception ex)
            {
                lm.Write("DataSetManager/LoadTodaysDataSet:  " + ex.Message);
            }
        }
        /* if (status.Equals("On Order"))
                GetPONumber(reqItem);
         */

        public void GetPONumber(int reqItemID)
        {
            string reqno_po = "";
            string key = "";
            DataSet dsPO = new DataSet();
            ODMRequest Request = new ODMRequest();
            Request.ConnectString = connectStrBIAdmin;   //connectStrHEMM;
            Request.CommandType = CommandType.Text;
            Request.Command =   "SELECT PO_NO, PO_LINE.ITEM_ID " +
                                "FROM [h-hemm].dbo.PO_LINE " +
                                "JOIN [h-hemm].dbo.PO ON PO.PO_ID = PO_LINE.PO_ID " +
                                "JOIN [h-hemm].dbo.REQ ON REQ.REQ_NO = PO_LINE.REQ_NO " +
                                "WHERE PO_LINE.REQ_NO = " +
                                "RTRIM((SELECT REQ_ID FROM [h-hemm].dbo.REQ_ITEM WHERE REQ_ITEM_ID = " + reqItemID + ")) " +
                                "AND PO_LINE.ITEM_ID = (SELECT ITEM_ID FROM [h-hemm].dbo.REQ_ITEM WHERE REQ_ITEM_ID = " + reqItemID + ")";

            /*
             * "SELECT VPO.PO_NO  " +
                               "FROM [h-hemm].dbo.v_hmcmm_Purchase_Orders VPO " +
                               "JOIN [h-hemm].dbo.REQ_ITEM ON REQ_ITEM.ITEM_ID = VPO.ITEM_ID WHERE REQ_NO = " +
                               "(SELECT REQ_NO FROM [h-hemm].dbo.REQ JOIN [h-hemm].dbo.REQ_ITEM ON REQ.REQ_ID = REQ_ITEM.REQ_ID WHERE REQ_ITEM_ID = " + reqItemID + ") " +
                               "AND REQ_ITEM_ID  = " + reqItemID;
             * */

            try
            {
                dsPO = ODMDataSetFactory.ExecuteDataSetBuild(ref Request);
                foreach (DataRow drow in dsPO.Tables[0].Rows)
                {
                    if(!reqItemPONO.ContainsKey(reqItemID))
                        reqItemPONO.Add(reqItemID, drow[0].ToString().Trim());
                }
                //  if (debug) lm.Write("DataSetManager/GetPONumber: reqNo_PO Count = " + reqNo_PO.Rows.Count);
            }
            catch (Exception ex)
            {
                lm.Write("DataSetManager/GetPONumber:  " + ex.Message);
            }
        }

        public void UpdateReqItems()
        {
            // private Hashtable itemsThatChanged 
            int item = 0;
            string status = "";

            if (debug) lm.Write("DataSetManager/UpdateReqItems");
            ODMRequest Request = new ODMRequest();
            Request.ConnectString = connectStrBIAdmin;
            Request.CommandType = CommandType.Text;
            foreach (object key in itemsThatChanged.Keys)
            {         
                try
                {
                    item = Convert.ToInt32(key);
                    status = itemsThatChanged[item].ToString();
                    Request.Command = "Execute ('" + BuildReqItemUpdateQuery(item,status) + "')";
                    ODMDataSetFactory.ExecuteDataWriter(ref Request);                    
                }
                catch (Exception ex)
                {
                    lm.Write("DataSetManager/UpdateReqItems:  " + ex.Message);
                }
            }
        }

        public void DeleteKilledComplete()
        { //remove KILLED and COMPLETE reqItems from the hmcmm_ReqItemStatus table
            int item = 0;
            string status = "";
            ODMRequest Request = new ODMRequest();
            Request.ConnectString = connectStrBIAdmin;
            Request.CommandType = CommandType.Text;
            Request.Command = BuildReqItemDeleteQuery();
            try
            {
                GetKilledCompleteCount(); //the "before" count
                ODMDataSetFactory.ExecuteNonQuery(ref Request);
                lm.Write("DataSetManager/DeleteKilledComplete:  ENQ Complete");
                GetKilledCompleteCount(); //the "after" count
            }
            catch (Exception ex)
            {
                lm.Write("DataSetManager/DeleteKilledComplete:  " + ex.Message);
            }          
        }        

        public void InsertTodaysList()
        {
            try
            {
              // CheckForSingleQuotes( "3,3'-DIAMINOBENZIDINE - 5GM'");         //for debug        
                ODMRequest Request = new ODMRequest();
                Request.ConnectString = connectStrBIAdmin;
                Request.CommandType = CommandType.Text;
                foreach (DataRow drow in dsCurrent.Tables[0].Rows)
                {
                    bool goodToGo = GetExistingRecordCount(Convert.ToInt32(drow[3]), Convert.ToInt32(drow[1]));
                    if (goodToGo)
                    {
                        Request.Command = BuildReqItemInsertQuery(drow);
                        //  lm.Write("DataSetManager/InsertTodaysList:  " + Environment.NewLine + Request.Command);
                        ODMDataSetFactory.ExecuteNonQuery(ref Request);
                    }
                }
            }
            catch (Exception ex)
            {
                lm.Write("DataSetManager/InsertTodaysList:  " + ex.Message);
            }
        }

        public void LoadYesterdayList()
        {
            //select req_item_id and req_item_stat and put into to a hashtable
            ODMRequest Request = new ODMRequest();
            Request.ConnectString = connectStrBIAdmin;
            Request.CommandType = CommandType.Text;
            Request.Command = "Execute ('" + BuildYesterdayQuery() + "')";

            if (debug)
                lm.Write("DataSetManager/LoadYesterdayList:  " + Request.Command);
            try
            {
                dsYesterday = ODMDataSetFactory.ExecuteDataSetBuild(ref Request);
            }
            catch (Exception ex)
            {
                lm.Write("DataSetManager/LoadYesterdayList:  " + ex.Message);
            }
        }

        public void LoadCurrentChanges(string reqItemList)
        {
            ODMRequest Request = new ODMRequest();
            Request.ConnectString = connectStrBIAdmin;   //connectStrHEMM;
            Request.CommandType = CommandType.Text;
            Request.Command = "Execute ('" + BuildCurrentQuery(reqItemList) + "')";

            if (debug)
                lm.Write("DataSetManager/LoadCurrentChanges:  " + "BuildTodayQuery() - list of Req Items");   //Request.Command);
            try
            {
                dsCurrentChangeDates = ODMDataSetFactory.ExecuteDataSetBuild(ref Request);
            }
            catch (Exception ex)
            {
                lm.Write("DataSetManager/LoadCurrentChanges:  " + ex.Message);
            }
        }

        private string BuildTodayQuery()
        {
            string query = "SELECT DISTINCT " +
                           "RI.LINE_NO [REQ LINE NO], " +  //0
                           "REQ_ITEM_ID, " +                      //1
                           "ITEM.DESCR, " +                        //2
                           "REQ.REQ_ID, " +                        //3
                           "ITEM.ITEM_NO, " +                    //4
                           "CASE RI.STAT " +
                           "WHEN 1 THEN ''Open'' " +
                           "WHEN 2 THEN ''Pend Apvl'' " +
                           "WHEN 3 THEN ''Approved'' " +
                           "WHEN 4 THEN ''Removed'' " +
                           "WHEN 5 THEN ''Pending PO'' " +
                           "WHEN 6 THEN ''Open Stock Order'' " +
                           "WHEN 8 THEN ''Draft'' " +
                           "WHEN 9 THEN ''On Order'' " +
                           "WHEN 10 THEN ''Killed'' " +
                           "WHEN 11 THEN ''Complete'' " +
                           "WHEN 12 THEN ''Back Order'' " +
                           "WHEN 14 THEN ''On Order'' " +
                           "WHEN 24 THEN ''Pend Informational Apvl'' " +
                           "ELSE CAST(RI.STAT AS VARCHAR(2)) " +
                           "END [ITEM STAT],  " +                //5
                           "RI.STAT_CHG_DATE, " +   //6
                           "LOGIN_ID, " +                            //7
                           "REQ.REC_CREATE_DATE " + //8
                           "FROM [h-hemm].dbo.REQ_ITEM  RI  " +
                           "JOIN [h-hemm].dbo.REQ ON REQ.REQ_ID = RI.REQ_ID " +
                           "JOIN [h-hemm].dbo.ITEM ON ITEM.ITEM_ID = RI.ITEM_ID " +
                           "JOIN [h-hemm].dbo.USR ON USR.USR_ID = REQ.REC_CREATE_USR_ID " +
                           "WHERE REQ.REC_CREATE_DATE BETWEEN CONVERT(DATE,GETDATE()) AND CONVERT(DATE,GETDATE() + 1) " +
                           "AND LOGIN_ID <> ''iface'' " +
                          // "AND REQ_ITEM.STAT NOT IN (10,11)  " +                       //--10 Killed, 11 Complete
                           "order by 8,4,1 ";               //references param 7,3 and 0
            return query;
        }

        private string BuildYesterdayQuery()
        {
            string query =
                "SELECT REQ_ID,REQ_ITEM_ID, STAT_CHG_DATE,[REQ LINE NO],LOGIN_ID,DESCR,ITEM_NO " +
                "FROM hmcmm_ReqItemStatus Order by REQ_ID,[REQ LINE NO] DESC";
            return query;
        }
       
        private string BuildCurrentQuery(string reqItemList)
        {
            string query = 
            "SELECT " +
            "REQ_ITEM_ID, " +   //0
            "CASE REQ_ITEM.STAT " +
            "WHEN 1 THEN ''Open'' " +
            "WHEN 2 THEN ''Pend Apvl'' " +
            "WHEN 3 THEN ''Approved'' " +
            "WHEN 4 THEN ''Removed'' " +
            "WHEN 5 THEN ''Pending PO'' " +
            "WHEN 6 THEN ''Open Stock Order'' " +
            "WHEN 8 THEN ''Draft'' " +
            "WHEN 9 THEN ''On Order'' " +
            "WHEN 10 THEN ''Killed'' " +
            "WHEN 11 THEN ''Complete'' " +
            "WHEN 12 THEN ''Back Order'' " +
            "WHEN 14 THEN ''On Order'' " +
            "WHEN 15 THEN ''Auto PO'' " +             
             "WHEN 24 THEN ''Pend Informational Apvl'' " +
            "ELSE CAST(REQ_ITEM.STAT AS VARCHAR(2)) " +
            "END [ITEM STAT],  " +    //1
            "REQ_ITEM.STAT_CHG_DATE " +  //2
            "FROM [h-hemm].dbo.REQ_ITEM " +
            "WHERE REQ_ITEM.STAT = 9 " +        //only retrieves the On Order req items
            "AND REQ_ITEM_ID in ( " + reqItemList + 
            ")";
            return query;
        }

        private string BuildReqItemUpdateQuery(int reqItem, string status)
        {
            string query =
                         "UPDATE " +
                         "hmcmm_ReqItemStatus "+
                         "SET " +    
                         "STAT_CHG_DATE = ''"+ DateTime.Now + "'',"  +
                         "STAT = ''" + status + "'' " +
                         "WHERE REQ_ITEM_ID in (" + reqItem + ")";           
            return query;
        }

        private string BuildReqItemDeleteQuery()
        {
            string[] dt = DateTime.Now.ToString().Split(" ".ToCharArray());
            return "DELETE FROM dbo.hmcmm_ReqItemStatus WHERE STAT IN ('Killed', 'Complete', 'Removed') AND STAT_CHG_DATE < CONVERT(datetime, '" + dt[0] + "', 101)";          
        }

        private string BuildReqItemCountQuery(string dateToday)
        {
            return "SELECT COUNT(*) FROM dbo.hmcmm_ReqItemStatus WHERE STAT IN ('Killed', 'Complete','Removed') AND STAT_CHG_DATE < '" + dateToday + "'";
        }

        private string BuildReqItemInsertQuery(DataRow dRow)
        {
            string query = "";
                query = "INSERT INTO [dbo].[hmcmm_ReqItemStatus] VALUES(" +
                               dRow[0] + //REQ_LINE
                               "," + dRow[1] + //REQ_ITEM_ID
                               ",'" + CheckForSingleQuotes(dRow[2].ToString()) + //ITEM_DESC
                               "'," + dRow[3] + //REQ_ID
                               ",'" + dRow[4].ToString().Trim() + //ITEM_NO
                               "','" + dRow[5] + //STAT
                               "','" + dRow[6] + //STAT_CHG_DATE
                               "','" + dRow[7].ToString().Trim() + //LOGIN_ID
                               "','" + dRow[8].ToString().Trim() + //REC_CREATE_DATE 
                               "')";
            return query;
        }

        private bool GetExistingRecordCount(int REQ_ID, int RI_ID)
        {
            bool goodToGo = false;
            ArrayList  reqCount = new ArrayList();
            ODMRequest Request = new ODMRequest();
            Request.ConnectString = connectStrBIAdmin;
            Request.CommandType = CommandType.Text;
            Request.Command = "SELECT COUNT(*) FROM [dbo].[hmcmm_ReqItemStatus] WHERE REQ_ITEM_ID = " + RI_ID +
                              " AND REQ_ID = " + REQ_ID;
            reqCount = ODMDataSetFactory.ExecuteDataReader(Request);
            if (Convert.ToInt32(reqCount[0]) == 0)
                goodToGo = true;

            return goodToGo;
        }

        private void GetKilledCompleteCount()
        {
            ArrayList count = new ArrayList();
            ODMRequest Request = new ODMRequest();        
            string[] dtYesterday = DateTime.Now.AddDays(-1).ToString().Split(" ".ToCharArray()); //this prints yesterday's date
            string[] dtToday = DateTime.Now.ToString().Split(" ".ToCharArray()); //this is for the get COUNT(*) query.
            Request.ConnectString = connectStrBIAdmin;
            Request.CommandType = CommandType.Text;
            Request.Command = BuildReqItemCountQuery(dtToday[0]);
            try
            {
                count = ODMDataSetFactory.ExecuteDataReader(ref Request);
                //if (debug && count.Count > 0)                
                lm.Write("Killed&Complete Count for " + dtYesterday[0] + " : " + count[0]);
                
            }
            catch (Exception ex)
            {
                lm.Write("DataSetManager/GetKilledCompleteCount:  " + ex.Message);
            }
        }

        private string CheckForSingleQuotes(string desc)
        {
            string[] quote = desc.Split("'".ToCharArray());
            desc = "";
            for(int x = 0; x < quote.Length;x++)
            {               
                desc += quote[x] + "''";
            }
               desc = desc.Substring(0, desc.Length - 2);            
            return desc;
        }
    }
}
