using System;
using System.Collections;
using System.Net.Mail;
using System.IO;
using System.Data;
using KeyMaster;
using SpreadsheetLight;
using LogDefault;

namespace ReqItemStatus
{
    class OutputManager
    {
        #region ClassVariables
        private ArrayList reqItems = new ArrayList();
        private string userName = "";
        private string recipient = "";
        private string body = "";
        private string attachmentPath = "";
        private string extension = "";
        private string helpFile = "pmmhelp.txt";
        private string poNo = "";
        private string unameVarients = ""; 
        private bool debug = false;
        private char TAB = Convert.ToChar(9);
        private Hashtable itemReq = new Hashtable();
        private Hashtable itemReqLine = new Hashtable();
        private Hashtable itemsThatChanged = new Hashtable();
        private Hashtable itemItemNo = new Hashtable();
        private Hashtable itemDesc = new Hashtable();
        private Hashtable reqItemPO = new Hashtable();
        private LogManager lm = LogManager.GetInstance();
        private DataSetManager dsm = DataSetManager.GetInstance();
        private int fileCount = 1;
        #endregion
        #region parameters
        public ArrayList ReqItems
        {
            set { reqItems = value; }
        }
        public string UserName
        {
            set { userName = value; }
        }

        public string UnameVarients
        {
            set { unameVarients = value; }
        }

        public string AttachmentPath
        {
            set { attachmentPath = value; }
        }
        public Hashtable ItemReq
        {//reqItemID,reqNo
            set { itemReq = value; }
        }
        public Hashtable ItemReqLine
        {
            set { itemReqLine = value; }
        }
        public Hashtable ItemsThatChanged
        {
            set { itemsThatChanged = value; }
        }
        public Hashtable ItemItemNo
        {//reqItemID,itemNo
            set { itemItemNo = value; }
        }
        public Hashtable ItemDesc
        {
            set { itemDesc = value; }
        }
        public bool Debug
        {
            set { debug = value; }
        }
        #endregion

        public void SendOutput()
        {
            reqItemPO = (Hashtable)dsm.ReqItemPONO.Clone();
            FormatEmail();
            FormatAttachment();
            SendMail();
        }
       
        private void FormatEmail()
        {
            string desc = "";
            string status = "";
           // string poNo = "";

            if (debug) lm.Write("OutputManager/FormatEmail");

            body = "Below, and attached, is a list of your requisition items that have had a recent status change. " + Environment.NewLine + Environment.NewLine;
            body += "REQ NMBR" + TAB + "REQ LINE" + TAB + "ITEM#" + TAB + TAB + "NEW STATUS" + TAB +  TAB + "PO" + TAB + TAB + "DESCRIPTION" + 
                Environment.NewLine;

            if (!userName.Contains("@"))
                userName += "@uw.edu";
            recipient = userName;

            try
            {
                foreach(string reqitem in reqItems)
                {
                    Check_OnOrder_Status(Convert.ToInt32(reqitem));
                    desc = itemDesc[Convert.ToInt32(reqitem)].ToString();
                    desc = desc.Length > 40 ? desc.Substring(0, 40) : desc;
                    status = itemsThatChanged[Convert.ToInt32(reqitem)].ToString().Trim();
                    status += status == "Killed" ? "        " : "";
                    body += itemReq[Convert.ToInt32(reqitem)].ToString() + TAB +
                            itemReqLine[Convert.ToInt32(reqitem)].ToString() + TAB + TAB +
                            CheckNonStock(itemItemNo[Convert.ToInt32(reqitem)].ToString().Trim()) + TAB + TAB +
                            status + TAB + TAB +
                            poNo + TAB + TAB +
                            desc + Environment.NewLine;                  
                }
            }
            catch (Exception ex)
            {
                lm.Write("OutputManager/FormatEmail:  " + ex.Message);
            }
        }

        private string CheckNonStock(string itmNo)
        {
            if (itmNo.Contains("~"))
                itmNo = "Non Catalog";
            return itmNo;
        }

        private void FormatAttachment()
        {            
            SLDocument sld = new SLDocument();            
            string mssg = "There's Nothing to Export";
            int reqItemID = 0;
            int colNo = 0;
            int rowNo = 1;
            //poNo = "pending";
            try
            {                //set the col headers
                sld.SetCellValue(rowNo,++colNo, "REQ NMBR" );
                sld.SetCellValue(rowNo,++colNo, "REQ LINE" );
                sld.SetCellValue(rowNo,++colNo, "ITEM#" );
                sld.SetCellValue(rowNo,++colNo, "NEW STATUS" );
                sld.SetCellValue(rowNo, ++colNo, "PO");
                sld.SetCellValue(rowNo,++colNo, "DESCRIPTION" );                

                foreach (string reqitem in reqItems)
                {
                    reqItemID = Convert.ToInt32(reqitem);
                    colNo = 0;
                    rowNo++;
                    poNo = dsm.ReqItemPONO.ContainsKey(reqItemID) ? dsm.ReqItemPONO[reqItemID].ToString() : "";
                    sld.SetCellValue(rowNo, ++colNo, itemReq[reqItemID].ToString());
                    sld.SetCellValue(rowNo, ++colNo, itemReqLine[reqItemID].ToString());
                    sld.SetCellValue(rowNo, ++colNo, CheckNonStock(itemItemNo[reqItemID].ToString().Trim()));
                    sld.SetCellValue(rowNo, ++colNo, itemsThatChanged[reqItemID].ToString().Trim());
                    sld.SetCellValue(rowNo, ++colNo, poNo);
                    sld.SetCellValue(rowNo, ++colNo, itemDesc[Convert.ToInt32(reqItemID)].ToString().Trim());                                      
                }
                extension = "Req Item Status" + fileCount++ + ".xlsx";
                sld.SaveAs(attachmentPath + extension);
            }
            catch (IndexOutOfRangeException ex)
            {
                lm.Write("OutputManager/FormatAttachment:  IOOR Exception  " + ex.Message);
            }
            catch (Exception ex)
            {
                lm.Write("OutputManager/FormatAttachment:  Exception  " + ex.Message);
            }
        }

        private void SendMail()
        {
           // if (debug) lm.Write("OutputManager/SendMail");
            try
            {
                string logMssg = "";
                bool mailError = false;                
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.uw.edu");
                mail.From = new MailAddress("pmmhelp@uw.edu");

                if (debug)
                    mail.To.Add("dlrjones@uw.edu");
                else
                {
                    mail.To.Add(recipient);
                }
                mail.Subject = "Status Change in Requisitions";
                if (debug)
                    mail.Subject += "       " + recipient;
                mail.Body = body + 
                            Environment.NewLine + Environment.NewLine +
                            "Thanks," +
                            Environment.NewLine +
                            Environment.NewLine +
                            "PMMHelp" + Environment.NewLine +
                            "UW Medicine Harborview Medical Center" + Environment.NewLine +
                            "Supply Chain Management Informatics" + Environment.NewLine +
                            "206-598-0044" + Environment.NewLine +
                            "pmmhelp@uw.edu";
                mail.ReplyToList.Add("pmmhelp@uw.edu");

                Attachment attachment;
                attachment = new System.Net.Mail.Attachment(attachmentPath + extension);
                mail.Attachments.Add(attachment);

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("pmmhelp", GetKey());
                SmtpServer.EnableSsl = true;

                try
                { //debug=true sends all emails to dlrjones@uw.edu
                    if (!debug)
                        SmtpServer.Send(mail);

                    logMssg = "OutputManager/SendMail:  Sent To  " + mail.To;
                    if (debug)
                    {
                        logMssg += "       (for " + recipient + ")";
                        SmtpServer.Send(mail); //comment this out to prevent emails going to dlrjones while debug = true
                    }
                }
                catch (SmtpException ex)
                {
                    lm.Write("SendMail SmtpException:  " + ex.Message + Environment.NewLine + ex.InnerException);
                    mailError = true;
                }
                catch (Exception ex)
                {
                    lm.Write("SendMail Error " + ex.Message+ Environment.NewLine + ex.InnerException);
                    mailError = true;
                }
                if (mailError)
                {//sometimes SendMail errors out with a message like "The message or signature supplied for verification has been altered"
                    //or "The buffers supplied to a function was too small". Sending it a second time sometimes works.
                    try
                    {
                        SmtpServer.Send(mail);
                        lm.Write("SendMail mail resent: " + mail.To.ToString() + "       (for " + recipient + ")");
                    }
                    catch (Exception ex)
                    {
                        lm.Write("SendMail mailError 2  " + ex.Message + Environment.NewLine + ex.InnerException);
                    }
                }
                lm.Write(logMssg);
            }
            catch (Exception ex)
            {
                string mssg = ex.Message;
                lm.Write("OutputManager/SendMail: Exception    " + mssg);
            }
        }

        public string GetKey()
        {
            string[] key = File.ReadAllLines(attachmentPath + helpFile);
            return StringCipher.Decrypt(key[0],"pmmhelp");                  
            }

        private void Check_OnOrder_Status(int reqItem)
        {
            string stat = itemsThatChanged[reqItem].ToString().Trim();
            poNo = "";
            try
            {                
                if (stat == "On Order")
                {                    
                    if(reqItemPO.ContainsKey(reqItem))
                    {                      
                        poNo = reqItemPO[reqItem].ToString();
                    }                           
                }                
            }
            catch (Exception ex)
            {
                lm.Write("OutputManager/CheckOnOrderStatus:  " + ex.Message + Environment.NewLine + ex.InnerException);
            }
        }
    
    }
}
