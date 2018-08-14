using System;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.IO;

namespace ReqItemStatus
{    
    class ProcessManager
    {
        #region Class Variables & Parameters
        private DataSet dsYesterday = new DataSet();
        private DataSet dsCurrentChangeDates;        
        private Hashtable itemChangeDate = new Hashtable();     //reqItemID,chngDate
        private Hashtable itemReq = new Hashtable();            //reqItemID,reqNo
        private Hashtable itemReqLine = new Hashtable();        //reqItemID,reqLine
        private Hashtable itemsThatChanged = new Hashtable();   //reqItemID,status
        private Hashtable itemItemNo = new Hashtable();         //reqItemID,itemNo
        private Hashtable itemDesc = new Hashtable();           //reqItemID,desc
        private Hashtable itemLogin = new Hashtable();          //reqItemID,login
        private Hashtable userItems = new Hashtable();
        private LogManager lm = LogManager.GetInstance();
        private DataSetManager dsm = DataSetManager.GetInstance();
        private static NameValueCollection ConfigData = null;
        private string reqItemIDs = "";
        private bool debug = false;       

        class UserNameVariant
        {
            //takes in the path to the text file containing the names of users who have an email id different from their AMC id.
            //each entry looks like this - mdanna|dannam
            Hashtable userItems = new Hashtable();
            private string unamePath = "";
            private LogManager lm = LogManager.GetInstance();

            public string UnamePath
            {
                set { unamePath = value; }
            }
            public Hashtable UserItems
            {
                get { return userItems; }
                set
                {
                    userItems = value; 
                    GetUserNameVariant();
                }
            }
           
            private void GetUserNameVariant()
            {
                string[] users = File.ReadAllLines(unamePath);
                ArrayList tmpValu = new ArrayList();
                string[] user;

                foreach (string name in users)
                {
                    user = name.Split("|".ToCharArray());
                    if (user.Length > 0)
                    {
                        if (user.Length > 1)
                        {
                            //change the userItems entry for users that have a different email
                            if (userItems.ContainsKey(user[0]))
                            {
                                try
                                {
                                    tmpValu = (ArrayList)userItems[user[0]];
                                    userItems.Remove(user[0]);
                                    userItems.Add(user[1], tmpValu);
                                }
                                catch (Exception ex)
                                {
                                    lm.LogEntry("ProcessManager/GetUserNameVariant:  " + ex.Message);
                                }
                            }
                        }
                    }
                }
            }
  //          WHERE LOGIN_ID IN ('mharrington','sbuckingham','mdanna','bsimmons','ayyoub','fyoshioka',
  //'jvarghese','mvasiliades','bwalker','dlam','sgoss','enewcombe','aclouse','jrudd',
  //'kburkette','ewardenburg','mofo','ccastillo')
        }

        public DataSet DSYesterday
        {
            get { return dsYesterday; }
            set { dsYesterday = value; }
        }

        public bool Debug
        {
            set { debug = value; }
        }
 #endregion

       public  ProcessManager()
        {
            ConfigData = (NameValueCollection)ConfigurationSettings.GetConfig("appSettings");
        }

        public void Begin()
        {
            LoadYesterdayHash();
            LoadChangedDates();
            CheckStatus();  //Fills the itemsThatChanged Hashtable
            OrganizeByLogin(); //create and fill the userItems Hashtable
            PrepareOutput(); //gets instance of OutputManager
            UpdateStatus();
            StoreNewReqs();    
        }

        private void LoadYesterdayHash()
        {
            /*
             * [0]-REQ_ID,  [1]-REQ_ITEM_ID,  [2]-STAT_CHNG_DATE,  [3]-REQ_LINE,  [4]-LOGIN_ID [5]-ITEM_DESC, [6]-ITEM_NO
             * */

            if (debug)  lm.LogEntry("ProcessManager/LoadYesterdayHash:");
            int reqItemID = 0;
            int reqNo = 0;
            int reqLine = 0;
            string login = "";
            string desc = "";
            string itemNo = "";
            DateTime chngDate = new DateTime();
            dsYesterday = dsm.DsYesterday;  //initialized in Program/LoadData
            
            try
            {
                foreach (DataRow drow in dsYesterday.Tables[0].Rows)
                {
                    if (drow[0].ToString().Trim().Length > 0)
                        reqNo = Convert.ToInt32(drow[0]);
                    if (drow[1].ToString().Trim().Length > 0)
                        reqItemID = Convert.ToInt32(drow[1]);
                    if (drow[2].ToString().Trim().Length > 0)
                        chngDate = Convert.ToDateTime(drow[2]);
                    if (drow[3].ToString().Trim().Length > 0)
                        reqLine = Convert.ToInt32(drow[3]);
                    if (drow[4].ToString().Trim().Length > 0)
                        login = drow[4].ToString();
                    if (drow[5].ToString().Trim().Length > 0)
                        desc = drow[5].ToString();
                    if (drow[6].ToString().Trim().Length > 0)
                        itemNo = drow[6].ToString();

                    if (!itemChangeDate.ContainsKey(reqItemID))
                        itemChangeDate.Add(reqItemID, chngDate);
                    if (!itemReq.ContainsKey(reqItemID))
                        itemReq.Add(reqItemID,reqNo);
                    if (!itemReqLine.ContainsKey(reqItemID))
                        itemReqLine.Add(reqItemID,reqLine);
                    if (!itemLogin.ContainsKey(reqItemID))
                        itemLogin.Add(reqItemID,login);
                    if (!itemDesc.ContainsKey(reqItemID))
                        itemDesc.Add(reqItemID,desc);
                    if (!itemItemNo.ContainsKey(reqItemID))
                        itemItemNo.Add(reqItemID,itemNo);

                    reqItemIDs += reqItemIDs.Length > 0 ? "," + reqItemID.ToString() : reqItemID.ToString();

                }
            }
            catch (Exception ex)
            {
                lm.LogEntry("ProcessManager/LoadYesterdayHash:  " + ex.Message);
            }
        }

        private void LoadChangedDates()
        {//use the list of reqItemIDs from the ReqItemStatus table and loaded in the LoadYesterdayHash method
            //dsm will give back the current status of that line item
            if (debug)  lm.LogEntry("ProcessManager/LoadChangedDates");
            dsm.DSCurrentChangeDates = dsCurrentChangeDates;
            dsm.Debug = debug;
            dsm.LoadCurrentChanges(reqItemIDs); 
            dsCurrentChangeDates = dsm.DSCurrentChangeDates;
        }

        private void CheckStatus()
        {
            /*
             * Fills the itemsThatChanged Hashtable:
             * dsCurrentChangeDates holds the current status of those req_item_id's that are in the uwm_BIAdmin/hmcmm_ReqItemStatus table
             * iterate through the list comparing the status found with the status from yesterday (found in the Hashtable itemChangeDate)
             */
            if (debug)  lm.LogEntry("ProcessManager/CheckStatus");
            int reqItemID = 0;
            DateTime chngDate = new DateTime();
            string newStatus = "";

            //dsCurrentChangeDates =  [0]-REQ_ITEM_ID,  [1]-STATUS,  [2]-STAT_CHNG_DATE
            try
            {
                if (dsCurrentChangeDates.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow drow in dsCurrentChangeDates.Tables[0].Rows)
                    {
                        if (drow[1] != DBNull.Value)
                            newStatus = drow[1].ToString();

                        if (drow[0] != DBNull.Value && drow[2] != DBNull.Value)
                        {
                            reqItemID = Convert.ToInt32(drow[0]);
                            chngDate = Convert.ToDateTime(drow[2]);
                       //     if (debug) lm.LogEntry("ProcessManager/CheckStatus reqItemID = " + reqItemID + "  STAT = " + newStatus);
                       //     if (newStatus.Equals("On Order"))
                       //         dsm.GetPONumber(reqItemID);
                            if (chngDate > Convert.ToDateTime(itemChangeDate[reqItemID]))
                            {
                                itemsThatChanged.Add(reqItemID, drow[1].ToString());
                                if (newStatus.Equals("On Order"))
                                    dsm.GetPONumber(reqItemID);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                lm.LogEntry("ProcessManager/CheckStatus:  " + ex.Message);
            }
        }

        private void OrganizeByLogin()
        {/*
          * Fills the itemsPerLogin Hashtable:
          * this inverts the itemLogin Hashtable (populated in LoadYesterdayHash())
          * The user's login id is the key and the value is an ArrayList containing all
          * of the reqItemID's associated that user. This is later used to consolidate
          * all req lines onto one email to the user.
          * */
            if (debug)  lm.LogEntry("ProcessManager/OrganizeByLogin");
          //  ArrayList statusInfo = new ArrayList();
            ArrayList itemsPerLogin = new ArrayList();
            ArrayList itemsNotInChangedList = new ArrayList();
            Hashtable cloneLoginHash = new Hashtable();
            string login = "";
            string valu = "";
            bool changeValidated = false;
            int changeCount = itemLogin.Count;
         //   int notUsedCount = 0;

            cloneLoginHash = Clone(itemLogin);
            try
            {
                while (changeCount > 0)
                {//the while loop goes through  each user in the itemLogin Hashtable 
                    foreach (object key in cloneLoginHash.Keys)
                    { //this ParallelLoopResult gets one user name and finds all instances of it in the itemLogin Hashtable
                        valu = cloneLoginHash[key].ToString().Trim();
                        if (login.Length == 0) //identify the current user
                            login = valu;

                        //ignore all but the current user for this iteration of the loop
                        if (login == cloneLoginHash[key].ToString().Trim())
                        {
                            if (itemsThatChanged.ContainsKey(key))
                            { // Add to the ArrayList of yesterday's items whose status has changed 
                                itemsPerLogin.Add(key.ToString());
                                changeValidated = true;
                            }
                            else if (!(itemsNotInChangedList.Contains(key)))
                                //Add to the ArrayList of items with no change in status
                                itemsNotInChangedList.Add(key);
                        }                     
                    }

                    if (!userItems.ContainsKey(login) && changeValidated)
                        userItems.Add(login, itemsPerLogin.Clone()); //user login id & arraylist of their reqItemID's  --
                    foreach (object item in itemsPerLogin)
                    {//remove the 'status changed' reqItemID's for the specific user
                        cloneLoginHash.Remove(Convert.ToInt32(item));
                        changeCount--;
                    }
                    foreach (object item in itemsNotInChangedList)
                    {//remove the 'no change' reqItemID's for the specific user
                        cloneLoginHash.Remove(Convert.ToInt32(item));
                        changeCount--;
                    }

                    login = "";
                    changeValidated = false;
                    itemsPerLogin. Clear();
                    itemsNotInChangedList.Clear();
                }
            }
            catch (Exception ex)
            {
                lm.LogEntry("ProcessManager/CheckStatus:  " + ex.Message);
            }
        }

        private void PrepareOutput()
        {
            OutputManager om = new OutputManager();
            om.Debug = debug;
            om.AttachmentPath = ConfigData.Get("attachmentPath");
            om.ItemReq = itemReq;
            om.ItemReqLine = itemReqLine;
            om.ItemsThatChanged = itemsThatChanged;
            om.ItemItemNo = itemItemNo;
            om.ItemDesc = itemDesc;
            GetUserNameVariant();
            int itemID = 0;
            try
            {
                foreach (DictionaryEntry dictionaryEntry in userItems)
                {                
                    om.UserName = dictionaryEntry.Key.ToString();
                    om.ReqItems = new ArrayList((ArrayList)dictionaryEntry.Value);
                    om.SendOutput();
                }
            }
            catch (Exception ex)
            {
                lm.LogEntry("ProcessManager/PrepareOutput:  ERROR:  " + ex.Message);
            }
        }

        //changes are made to the userItems list to replace the user names which AREN'T also the email name        
        private void GetUserNameVariant()
        {
            UserNameVariant unv = new UserNameVariant();
            unv.UnamePath = ConfigData.Get("unameVariantList");
            unv.UserItems = userItems;
            userItems = unv.UserItems;
        }

        private void UpdateStatus()
        {
            if (debug) lm.LogEntry("ProcessManager/UpdateStatus");
            dsm.ItemsThatChanged = Clone(itemsThatChanged);
            dsm.UpdateReqItems();
        }

        private void StoreNewReqs()
        {
            if (debug) lm.LogEntry("ProcessManager/StoreNewReqs");
            dsm.InsertTodaysList();
        }      
       
        private Hashtable Clone(Hashtable htInput)
        {
            if (debug) lm.LogEntry("ProcessManager/Clone");
            Hashtable ht = new Hashtable();

            foreach (DictionaryEntry dictionaryEntry in htInput)
            {
                if (dictionaryEntry.Value is string)
                {
                    ht.Add(dictionaryEntry.Key, new string(dictionaryEntry.Value.ToString().ToCharArray()));
                }
                else if (dictionaryEntry.Value is Hashtable)
                {
                    ht.Add(dictionaryEntry.Key, Clone((Hashtable)dictionaryEntry.Value));
                }
                else if (dictionaryEntry.Value is ArrayList)
                {
                    ht.Add(dictionaryEntry.Key, new ArrayList((ArrayList)dictionaryEntry.Value));
                }
            }
            return ht;
        }
    }
}
