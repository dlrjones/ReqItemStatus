// Email Password
/*
The password for the SendMail portion of this app is stored in the file [attachmentPath]\pmmhelp.txt (find attachmentPath in the config file).
The referenced library KeyMaster is used to decrypt the password at run time. There is another app called EncryptAndHash that you
can use to change the password when that becomes necessary. The key is pmmhelp and the path to EncryptAndHash is 
\\Lapis\h_purchasing$\Purchasing\PMM IS data\HEMM Apps\Executables\.
*/




//license for SpreadSheetLight
/*
 * Copyright (c) 2011 Vincent Tan Wai Lip

Permission is hereby granted, free of charge, to any person obtaining a copy of this software
and associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial
portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */


 //	SQL
 /*
 This will return specific PO information given the req_item_id (as in DataSetManager/GetPONumber):
 
 SELECT VPO.REQ_NO, VPO.PO_NO, VPO.ITEM_NO,VPO.LINE_NO 
FROM v_hmcmm_Purchase_Orders VPO
JOIN REQ_ITEM ON REQ_ITEM.ITEM_ID = VPO.ITEM_ID
where REQ_NO = (SELECT REQ_NO FROM REQ JOIN REQ_ITEM ON REQ.REQ_ID = REQ_ITEM.REQ_ID WHERE REQ_ITEM_ID = [Req_Item_ID])
AND REQ_ITEM_ID = [Req_Item_ID]
 
 */