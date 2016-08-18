using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.Web.Caching;
using System.Web.Services;
using System.Xml;
using DAT_HHD;
using SAPbobsCOM;
namespace WebService
{
    /// <summary>
    /// Summary description for Common
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class Common : BaseData
    {       
        Company _company;

        #region ---- SAP Connection ----
        
        [WebMethod(EnableSession = true)]
        public string ConnectTest()
        {
            string sid = "";
            _company = new Company();
            _company.LicenseServer = ConfigurationManager.AppSettings["SAPlicense"];
            _company.Server = ConfigurationManager.AppSettings["SAPServer"];
            _company.CompanyDB = ConfigurationManager.AppSettings["CompanyDB"];
            _company.UserName = ConfigurationManager.AppSettings["SAPuserName"];
            _company.Password = ConfigurationManager.AppSettings["SAPpassword"];
            _company.language = BoSuppLangs.ln_English;
            _company.UseTrusted = Convert.ToBoolean(ConfigurationManager.AppSettings["SAPtursted"]);
            _company.DbUserName = ConfigurationManager.AppSettings["DbUserName"];
            _company.DbPassword = ConfigurationManager.AppSettings["DbPassword"];
            _company.DbServerType = BoDataServerTypes.dst_MSSQL2012;
            if (0 == _company.Connect())
            {
                int version = _company.Version;
                Guid id = Guid.NewGuid();
                sid = id.ToString();
                DateTime dtExpiration = DateTime.UtcNow.AddDays(2);
                HttpContext.Current.Cache.Insert(sid, (object)_company, null, dtExpiration, Cache.NoSlidingExpiration);

            }
            else
            {
                int errNo;
                string errMsg;
                _company.GetLastError(out errNo, out errMsg);
                return errMsg;
            }

            return sid;
        }
        
        #endregion

        #region ---- Common Method ----

        private string RemoveColon(string message)
        {
            string result = string.Empty;
            if (message.Length > 0)
            {
                if ((message[message.Length - 1].ToString() == ":"))
                {
                    result = message.Substring(0, message.Length - 1);
                }
                else
                {
                    result = message;
                }
            }
            return result;
        }

        private string AlertNotification(string guid, string subject, string content, string docNum, string docEntry, string docObject, string module)
        {
            DataSet _dataSet = GetEmployeeList();
            string errMsg = "";
            List<string> recipients = new List<string>();
            if (module == "Sales")
            {
                if (_dataSet.Tables.Count > 0 && _dataSet.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dRow in _dataSet.Tables[0].Rows)
                    {
                        recipients.Add(dRow["UserName"].ToString());
                    }
                }
            }
            else
            {
                if (_dataSet.Tables.Count > 0 && _dataSet.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dRow in _dataSet.Tables[0].Rows)
                    {
                        if (dRow["DeptName"].ToString().Trim() != "CS")
                        {
                            recipients.Add(dRow["UserName"].ToString());
                        }
                    }
                }
            }
            if (recipients.Count > 0)
            {
                errMsg = SendAlert(guid, subject, content, docNum, docEntry, docObject, recipients);
            }
            return errMsg;

        }

        private DataSet GetEmployeeList()
        {
            DataSet _sapDataSet = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select empID 'EmpId', dept 'DeptId', (select Name from OUDP where Code=dept) AS 'DeptName',USER_CODE 'UserName', SUPERUSER 'SuperUser', Active from OHEM T0 Inner Join OUSR T1 ON T0.userId=T1.USERID Where Active='Y' AND (dept in((select Code from OUDP where Name='CS'),(select Code from OUDP where Name='LOGISTICS'),(select Code from OUDP where Name='PRODUCTION')) OR SUPERUSER='Y')";
            try
            {
                SqlDataAdapter _sapDataAdapter = new SqlDataAdapter(sqlcmd);
                _sapDataAdapter.Fill(_sapDataSet);
            }
            catch (Exception SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return _sapDataSet;
        }

        private string TriggerAlertMessage(string guid, string docEntry, string objType, string module, string docInfo, string content)
        {
            string result = null;
            string docNumber = null;
            DataSet _dSet = null;
            if (objType == "59")
            {
                _dSet = GetDraftDataByDocumentNo("OIGN", docEntry);
                if (_dSet.Tables.Count > 0 && _dSet.Tables[0].Rows.Count > 0)
                {
                    docNumber = _dSet.Tables[0].Rows[0]["DocNum"].ToString();
                }
            }
            else if (objType == "60")
            {
                _dSet = GetDraftDataByDocumentNo("OIGE", docEntry);
                if (_dSet.Tables.Count > 0 && _dSet.Tables[0].Rows.Count > 0)
                {
                    docNumber = _dSet.Tables[0].Rows[0]["DocNum"].ToString();
                }
            }
            else if (objType == "112")
            {
                _dSet = GetDraftDataByDocumentNo("ODRF", docEntry);
                if (_dSet.Tables.Count > 0 && _dSet.Tables[0].Rows.Count > 0)
                {
                    docNumber = _dSet.Tables[0].Rows[0]["DocNum"].ToString();
                }
            }
            result = AlertNotification(guid, docInfo, content, docNumber, docEntry, objType, module);
            return result;
        }

        private DataSet GetDraftDataByDocumentNo(string docTable, string docEntry)
        {
            DataSet _sapDataSet = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "Select DocEntry,DocNum from " + docTable + " where DocStatus='O' AND DocEntry=" + docEntry + "";
            try
            {
                SqlDataAdapter _sapDataAdapter = new SqlDataAdapter(sqlcmd);
                _sapDataAdapter.Fill(_sapDataSet);
            }
            catch (Exception SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return _sapDataSet;
        }

        private string SendAlert(string guid, string subject, string content, string docNum, string docEntry, string docObject, List<string> Recipients)
        {
            SAPbobsCOM.CompanyService oCmpSrv = null;
            MessagesService oMessageService = null;
            SAPbobsCOM.Message oMessage = null;
            MessageDataColumns pMessageDataColumns = null;
            MessageDataColumn pMessageDataColumn = null;
            MessageDataLines oLines = null;
            MessageDataLine oLine = null;
            RecipientCollection oRecipientCollection = null;
            Recipient oRecipient = null;
            try
            {
                //get company service
                oCmpSrv = ((Company)HttpContext.Current.Cache[guid]).GetCompanyService();

                //get message service
                oMessageService = (MessagesService)oCmpSrv.GetBusinessService(ServiceTypes.MessagesService);

                //get the data interface for the new message
                oMessage = (Message)oMessageService.GetDataInterface(MessagesServiceDataInterfaces.msdiMessage);

                //fill subject
                oMessage.Subject = docEntry + ". " + subject;

                //fill text
                oMessage.Text = content + ". Reference Document Number : " + docNum + ".";

                //Add Recipient
                oRecipientCollection = oMessage.RecipientCollection;
                
                foreach (string UserCode in Recipients)
                {
                    oRecipient = oRecipientCollection.Add();
                    oRecipient.SendInternal = BoYesNoEnum.tYES;
                    oRecipient.UserCode = UserCode;
                }

                //get columns data
                pMessageDataColumns = oMessage.MessageDataColumns;

                //get column
                pMessageDataColumn = pMessageDataColumns.Add();

                //set column name
                pMessageDataColumn.ColumnName = "Document Entry";
                pMessageDataColumn.Link = (int.Parse(docObject.Trim()) != -1 ? BoYesNoEnum.tYES : BoYesNoEnum.tNO);

                //get lines
                oLines = pMessageDataColumn.MessageDataLines;

                //add new line
                oLine = oLines.Add();

                //set the line value
                oLine.Value = docEntry;
                oLine.Object = docObject.Trim();
                oLine.ObjectKey = docEntry;

                //----------------------------------------

                pMessageDataColumn = pMessageDataColumns.Add();
                pMessageDataColumn.ColumnName = "Document No";
                pMessageDataColumn.Link = BoYesNoEnum.tNO;

                oLines = pMessageDataColumn.MessageDataLines;
                oLine = oLines.Add();
                oLine.Value = docNum;
                //----------------------------------------
                pMessageDataColumn = pMessageDataColumns.Add();
                pMessageDataColumn.ColumnName = "Document Info";
                pMessageDataColumn.Link = SAPbobsCOM.BoYesNoEnum.tNO;

                oLines = pMessageDataColumn.MessageDataLines;
                oLine = oLines.Add();
                oLine.Value = content;

                //send the message
                oMessage.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;
                oMessageService.SendMessage(oMessage);
                return "True";

            }
            catch (Exception SqlCee)
            {
                return SqlCee.Message;
            }
        }

        [WebMethod]
        public DataSet StockWithItem(string whsCode, string itemCode)
        {
            DataSet _sapDataSet = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select ItemCode, WhsCode, OnHand 'Quantity' from OITW where WhsCode ='" + whsCode + "' and ItemCode='" + itemCode + "'";
            try
            {
                SqlDataAdapter _sapDataAdapter = new SqlDataAdapter(sqlcmd);
                _sapDataAdapter.Fill(_sapDataSet);
            }
            catch (Exception SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return _sapDataSet;
        }

        [WebMethod]
        public DataSet StockWithBatch(string whsCode, string itemCode, string batchCode)
        {
            DataSet _sapDataSet = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "SELECT T0.WhsCode, T0.ItemCode, T1.DistNumber 'Batch', (T0.Quantity-T0.CommitQty) 'Quantity',T1.MnfDate,T1.Location FROM OBTQ T0 INNER JOIN OBTN T1 ON T0.MdAbsEntry = T1.AbsEntry WHERE T0.WhsCode ='" + whsCode + "' and T0.ItemCode='" + itemCode + "' and T1.DistNumber='" + batchCode + "'";
            try
            {
                SqlDataAdapter _sapDataAdapter = new SqlDataAdapter(sqlcmd);
                _sapDataAdapter.Fill(_sapDataSet);
            }
            catch (Exception SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return _sapDataSet;
        }

        [WebMethod]
        public DataSet GetUserById(string loginName, string loginPwd)
        {
            DataSet _sapDataSet = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "SELECT Code,Name,U_Loginname FROM [@ALE_TECSIAUSERS] WHERE U_Active='Y' AND U_LoginName='" + loginName + "' AND U_LoginPwd='" + loginPwd + "'";
            try
            {
                SqlDataAdapter _sapDataAdapter = new SqlDataAdapter(sqlcmd);
                _sapDataAdapter.Fill(_sapDataSet);
            }
            catch (Exception SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return _sapDataSet;
        }

        [WebMethod]
        public string GetItemByName(string itemName)
        {
            string ItemName = null;            
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "Select ItemCode from OITM Where ItemName='" + itemName + "'";
            try
            {
                SqlDataAdapter _sapDataAdapter = new SqlDataAdapter(sqlcmd);
                DataSet _sapDataSet = new DataSet();
                _sapDataAdapter.Fill(_sapDataSet);
                if (_sapDataSet.Tables.Count > 0 && _sapDataSet.Tables[0].Rows.Count > 0)
                {
                    ItemName = _sapDataSet.Tables[0].Rows[0]["ItemCode"].ToString();
                }
            }
            catch (Exception SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ItemName;
        }

        [WebMethod]
        public bool IsItemWithBatch(string itemCode)
        {
            bool exist = false;
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "Select CASE ManBtchNum WHEN 'Y' THEN 1 ELSE 0 END from OITM Where ItemCode='" + itemCode + "'";
            try
            {
                int count = (int)sqlcmd.ExecuteScalar();
                if (count == 1)
                    exist = true;
            }
            catch (Exception exception)
            {
                ExceptionLog(exception.ToString());

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return exist;
        }

        [WebMethod]
        public bool IsItemCodeExist(string itemCode)
        {
            bool exist = false;
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "Select count(ItemCode) from OITM Where ItemCode='" + itemCode + "'";
            try
            {
                int count = (int)sqlcmd.ExecuteScalar();
                if (count == 1)
                    exist = true;
            }
            catch (Exception exception)
            {
                ExceptionLog(exception.ToString());

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return exist;
        }

        [WebMethod]
        public DataSet GetWareHouses()
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            //this.sqlcmd.CommandText = "SELECT '-1' AS WhsCode, '-- Select Warehouse --' AS WhsName UNION ALL SELECT WhsCode, WhsName FROM OWHS";
            this.sqlcmd.CommandText = "SELECT WhsCode, WhsName FROM OWHS";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }

        [WebMethod]
        public DataSet GetItemCollections()
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "SELECT '-1' AS ItemCode, '-- Select Item --' AS ItemName UNION ALL SELECT ItemCode, ItemName FROM OITM Where validFor='Y'";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }

        #endregion

        #region ---- A/P Goods Receipt PO ----
        
        [WebMethod]
        public DataSet GetPOItem(string docNum)
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select LineNum,ItemCode, Dscription 'ItemName',null 'Location',WhsCode,null'BatchNo',OpenQty 'Quantity',null'ManuDate',null'BatchQty' from POR1 where OpenQty>0 AND DocEntry=(select DocEntry from OPOR where DocStatus='O' AND DocNum=" + docNum + ")";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }
               
        [WebMethod]
        public string GetDocEntryByDocNum(string docNum, string docObject)
        {
            string docEntry = null;
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select DocEntry from " + docObject + " where DocNum=" + docNum + "";
            try
            {
                SqlDataAdapter _sapDataAdapter = new SqlDataAdapter(sqlcmd);
                DataSet _sapDataSet = new DataSet();
                _sapDataAdapter.Fill(_sapDataSet);
                if (_sapDataSet.Tables.Count > 0 && _sapDataSet.Tables[0].Rows.Count > 0)
                {
                    docEntry = _sapDataSet.Tables[0].Rows[0]["DocEntry"].ToString();
                }
            }
            catch (Exception SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return docEntry;
        }

        /// <summary>
        /// This Method is used to create the Goods Receipt PO Draft
        /// </summary>
        /// <param name="guid">Globally Unique Identifier</param>
        /// <param name="docEntry"></param>
        /// <param name="xmlInner"></param>
        /// <returns></returns>
        
        [WebMethod(EnableSession = true)]
        public string GRPODraftCreate(string guid, int docEntry, string xmlInner)
        {
            string sErrMsg = "";
            int iErrorCode = 1;
            SAPbobsCOM.Documents objPurchaseOrder = null;
            SAPbobsCOM.Documents objPurchaseDelivery = null;
            try
            {

                string lineNumber = "";
                objPurchaseOrder = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlInner);
                XmlNodeList xmlNodeList = xmlDocument.SelectNodes("GRPO/LineNum");
                objPurchaseOrder.GetByKey(docEntry);
                objPurchaseDelivery = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                objPurchaseDelivery.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;

                #region --- Assign Header Properties ---

                objPurchaseDelivery.ContactPersonCode = objPurchaseOrder.ContactPersonCode;
                objPurchaseDelivery.SummeryType = objPurchaseOrder.SummeryType;
                objPurchaseDelivery.PayToCode = objPurchaseOrder.PayToCode;
                objPurchaseDelivery.PayToBankCountry = objPurchaseOrder.PayToBankCountry;
                objPurchaseDelivery.PayToBankCode = objPurchaseOrder.PayToBankCode;
                objPurchaseDelivery.PayToBankAccountNo = objPurchaseOrder.PayToBankAccountNo;
                objPurchaseDelivery.PayToBankBranch = objPurchaseOrder.PayToBankBranch;
                objPurchaseDelivery.PaymentBlockEntry = objPurchaseOrder.PaymentBlockEntry;
                objPurchaseDelivery.PaymentMethod = objPurchaseOrder.PaymentMethod;
                objPurchaseDelivery.Project = objPurchaseOrder.Project;
                objPurchaseDelivery.DocType = objPurchaseOrder.DocType;
                objPurchaseDelivery.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO;
                objPurchaseDelivery.DocDate = DateTime.Now;
                objPurchaseDelivery.RequriedDate = objPurchaseOrder.RequriedDate;
                objPurchaseDelivery.DocDueDate = DateTime.Now;
                objPurchaseDelivery.TaxDate = DateTime.Now;
                objPurchaseDelivery.CardCode = objPurchaseOrder.CardCode;
                objPurchaseDelivery.CardName = objPurchaseOrder.CardName;
                objPurchaseDelivery.Address = objPurchaseOrder.Address;
                objPurchaseDelivery.NumAtCard = objPurchaseOrder.NumAtCard;
                objPurchaseDelivery.Reference1 = objPurchaseOrder.Reference1;
                objPurchaseDelivery.Reference2 = objPurchaseOrder.Reference2;
                objPurchaseDelivery.Comments = "Based On Purchase Order " + objPurchaseOrder.DocNum + ".";
                objPurchaseDelivery.JournalMemo = objPurchaseOrder.JournalMemo;
                objPurchaseDelivery.GroupNumber = objPurchaseOrder.GroupNumber;
                objPurchaseDelivery.SalesPersonCode = objPurchaseOrder.SalesPersonCode;
                objPurchaseDelivery.TransportationCode = objPurchaseOrder.TransportationCode;
                objPurchaseDelivery.PartialSupply = SAPbobsCOM.BoYesNoEnum.tNO;
                objPurchaseDelivery.Confirmed = SAPbobsCOM.BoYesNoEnum.tYES;
                objPurchaseDelivery.UserFields.Fields.Item("U_DocTrans").Value = "HHD";

                #endregion

                #region --- Assign Line Properties ---

                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    XmlAttribute lineNum = xmlNode.Attributes["LineNum"];
                    if (lineNum != null)
                    {
                        lineNumber = lineNum.Value;
                    }
                    objPurchaseDelivery.Lines.ItemCode = xmlNode.SelectSingleNode("ItemCode").InnerText;
                    objPurchaseDelivery.Lines.AccountCode = objPurchaseOrder.Lines.AccountCode;
                    objPurchaseDelivery.Lines.ShipDate = objPurchaseOrder.Lines.ShipDate;
                    objPurchaseDelivery.Lines.BaseType = 22;
                    objPurchaseDelivery.Lines.BaseEntry = objPurchaseOrder.DocEntry;
                    objPurchaseDelivery.Lines.BaseLine = (Convert.ToInt32(lineNumber));
                    objPurchaseDelivery.Lines.Rate = objPurchaseOrder.Lines.Rate;
                    objPurchaseDelivery.Lines.DiscountPercent = objPurchaseOrder.Lines.DiscountPercent;
                    objPurchaseDelivery.Lines.VendorNum = objPurchaseOrder.Lines.VendorNum;
                    objPurchaseDelivery.Lines.SerialNum = objPurchaseOrder.Lines.SerialNum;
                    objPurchaseDelivery.Lines.WarehouseCode = xmlNode.SelectSingleNode("WhsCode").InnerText;
                    objPurchaseDelivery.Lines.SalesPersonCode = objPurchaseOrder.Lines.SalesPersonCode;
                    objPurchaseDelivery.Lines.CommisionPercent = objPurchaseOrder.Lines.CommisionPercent;
                    objPurchaseDelivery.Lines.AccountCode = objPurchaseOrder.Lines.AccountCode;
                    objPurchaseDelivery.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO;
                    objPurchaseDelivery.Lines.SupplierCatNum = objPurchaseOrder.Lines.SupplierCatNum;
                    objPurchaseDelivery.Lines.CostingCode = objPurchaseOrder.Lines.CostingCode;
                    objPurchaseDelivery.Lines.ProjectCode = objPurchaseOrder.Lines.ProjectCode;
                    objPurchaseDelivery.Lines.BarCode = objPurchaseOrder.Lines.BarCode;
                    objPurchaseDelivery.Lines.VatGroup = objPurchaseOrder.Lines.VatGroup;
                    objPurchaseDelivery.Lines.SWW = objPurchaseOrder.Lines.SWW;
                    objPurchaseDelivery.Lines.Address = objPurchaseOrder.Lines.Address;
                    objPurchaseDelivery.Lines.TaxCode = objPurchaseOrder.Lines.TaxCode;
                    objPurchaseDelivery.Lines.TaxType = SAPbobsCOM.BoTaxTypes.tt_Yes;
                    objPurchaseDelivery.Lines.BackOrder = SAPbobsCOM.BoYesNoEnum.tNO;
                    objPurchaseDelivery.Lines.FreeText = objPurchaseOrder.Lines.FreeText;
                    objPurchaseDelivery.Lines.Quantity = Convert.ToDouble(xmlNode.SelectSingleNode("Quantity").InnerText);
                    objPurchaseDelivery.Lines.UserFields.Fields.Item("U_TempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("Location").InnerText.Trim());
                    if (!string.IsNullOrEmpty(xmlNode.SelectSingleNode("BatchNo").InnerText))
                    {
                        string batchArray = xmlNode.SelectSingleNode("BatchNo").InnerText.TrimEnd(',');
                        string[] batchList = batchArray.Split(',');
                        for (int batch = 0; batch < batchList.Length; batch++)
                        {
                            char[] spliter = new char[] { '(', ')' };
                            string[] batchQty = batchList[batch].Split(spliter);
                            string[] manuDate = xmlNode.SelectSingleNode("ManuDate").InnerText.Split(',');
                            objPurchaseDelivery.Lines.BatchNumbers.BatchNumber = batchQty[0];
                            objPurchaseDelivery.Lines.BatchNumbers.ManufacturingDate = Convert.ToDateTime(manuDate[batch]);
                            objPurchaseDelivery.Lines.BatchNumbers.Quantity = double.Parse(batchQty[1]);
                            objPurchaseDelivery.Lines.BatchNumbers.Add();
                        }
                    }
                    objPurchaseDelivery.Lines.Add();
                }

                #endregion

                iErrorCode = objPurchaseDelivery.Add();
                ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetLastError(out iErrorCode, out sErrMsg);

            }
            catch (Exception ex)
            {
                sErrMsg = ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objPurchaseDelivery);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objPurchaseOrder);
            }
            if (iErrorCode == 0)
            {
                string currentdocNum = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey();
                string result = TriggerAlertMessage(guid, currentdocNum, "112", "Purchasing", "GRPO draft is created from HHD", "A/P - Goods Receipt PO draft is created from HHD");
                return result;
            }
            else
            {
                return sErrMsg;
            }
        }
        
        #endregion

        #region---- A/P Goods Return ---
        
        /// <summary>
        /// This Method is used to create the Goods Return Draft
        /// </summary>
        /// <param name="guid">Globally Unique Identifier</param>
        /// <param name="docEntry"></param>
        /// <param name="xmlInner"></param>
        /// <returns></returns>
        [WebMethod(EnableSession = true)]
        public string APGoodsReturnDraftCreate(string guid, int docEntry, string xmlInner)
        {
            string errMsg = "";
            int retValue = 1;
            SAPbobsCOM.Documents objPurchaseDelivery = null;
            SAPbobsCOM.Documents objPurchaseReturn = null;
            string lineNumber = "";
            try
            {
               
                objPurchaseDelivery = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlInner);
                XmlNodeList xmlNodeList = xmlDocument.SelectNodes("APGoodsreturn/LineNum");
                objPurchaseDelivery.GetByKey(docEntry);               
                objPurchaseReturn = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                objPurchaseReturn.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseReturns;

                #region --- Assign Header Properties ---

                objPurchaseReturn.ContactPersonCode = objPurchaseDelivery.ContactPersonCode;
                objPurchaseReturn.SummeryType = objPurchaseDelivery.SummeryType;
                objPurchaseReturn.PayToCode = objPurchaseDelivery.PayToCode;
                objPurchaseReturn.PayToBankCountry = objPurchaseDelivery.PayToBankCountry;
                objPurchaseReturn.PayToBankCode = objPurchaseDelivery.PayToBankCode;
                objPurchaseReturn.PayToBankAccountNo = objPurchaseDelivery.PayToBankAccountNo;
                objPurchaseReturn.PayToBankBranch = objPurchaseDelivery.PayToBankBranch;
                objPurchaseReturn.PaymentBlockEntry = objPurchaseDelivery.PaymentBlockEntry;
                objPurchaseReturn.PaymentMethod = objPurchaseDelivery.PaymentMethod;
                objPurchaseReturn.Project = objPurchaseDelivery.Project;
                objPurchaseReturn.DocType = objPurchaseDelivery.DocType;
                objPurchaseReturn.RequriedDate = objPurchaseDelivery.RequriedDate;
                objPurchaseReturn.Address = objPurchaseDelivery.Address;
                objPurchaseReturn.NumAtCard = objPurchaseDelivery.NumAtCard;
                objPurchaseReturn.CardCode = objPurchaseDelivery.CardCode;
                objPurchaseReturn.CardName = objPurchaseDelivery.CardName;
                objPurchaseReturn.DocDate = DateTime.Now;
                objPurchaseReturn.DocDueDate = DateTime.Now;
                objPurchaseReturn.TaxDate = DateTime.Now;
                objPurchaseReturn.Comments = objPurchaseDelivery.Comments + ". Based On Goods Receipt PO " + objPurchaseDelivery.DocNum + ".";
                objPurchaseReturn.SalesPersonCode = objPurchaseDelivery.SalesPersonCode;
                objPurchaseReturn.UserFields.Fields.Item("U_DocTrans").Value = "HHD";

                #endregion

                #region --- Assign Lines Properties ---

                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    XmlAttribute lID = xmlNode.Attributes["LineNum"];
                    if (lID != null)
                    {
                        lineNumber = lID.Value;
                    }
                    objPurchaseDelivery.Lines.SetCurrentLine((Convert.ToInt32(lineNumber)));
                    objPurchaseReturn.Lines.ItemCode = xmlNode.SelectSingleNode("ItemCode").InnerText;
                    objPurchaseReturn.Lines.ItemDescription = objPurchaseDelivery.Lines.ItemDescription;
                    objPurchaseReturn.Lines.ShipDate = objPurchaseDelivery.Lines.ShipDate;
                    objPurchaseReturn.Lines.BaseType = 20;
                    objPurchaseReturn.Lines.BaseEntry = objPurchaseDelivery.DocEntry;
                    objPurchaseReturn.Lines.BaseLine = objPurchaseDelivery.Lines.LineNum;
                    objPurchaseReturn.Lines.Rate = objPurchaseDelivery.Lines.Rate;
                    objPurchaseReturn.Lines.DiscountPercent = objPurchaseDelivery.Lines.DiscountPercent;
                    objPurchaseReturn.Lines.VendorNum = objPurchaseDelivery.Lines.VendorNum;
                    objPurchaseReturn.Lines.SerialNum = objPurchaseDelivery.Lines.SerialNum;
                    objPurchaseReturn.Lines.WarehouseCode = xmlNode.SelectSingleNode("WhsCode").InnerText;
                    objPurchaseReturn.Lines.AccountCode = objPurchaseDelivery.Lines.AccountCode;
                    objPurchaseReturn.Lines.Quantity = Convert.ToDouble(xmlNode.SelectSingleNode("Quantity").InnerText);
                    objPurchaseReturn.Lines.UserFields.Fields.Item("U_TempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("Location").InnerText.Trim());
                    if (!string.IsNullOrEmpty(xmlNode.SelectSingleNode("BatchNo").InnerText))
                    {
                        string batchArray = xmlNode.SelectSingleNode("BatchNo").InnerText.TrimEnd(',');
                        string[] batchList = batchArray.Split(',');
                        for (int batch = 0; batch < batchList.Length; batch++)
                        {
                            char[] spliter = new char[] { '(', ')' };
                            string[] batchQty = batchList[batch].Split(spliter);
                            objPurchaseReturn.Lines.BatchNumbers.BatchNumber = batchQty[0];
                            objPurchaseReturn.Lines.BatchNumbers.Quantity = double.Parse(batchQty[1]);
                            objPurchaseReturn.Lines.BatchNumbers.Add();
                        }
                    }
                    objPurchaseReturn.Lines.Add();
                }

                #endregion

                retValue = objPurchaseReturn.Add();
                ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetLastError(out retValue, out errMsg);
                if (retValue == -10)
                {
                    errMsg = "Items Have No Stock in the Warehouse";
                }
                if (retValue == 0)
                {
                    objPurchaseDelivery.Remove();
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objPurchaseReturn);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objPurchaseDelivery);
            }
            if (retValue == 0)
            {
                string currentdocNum = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey();
                string result = TriggerAlertMessage(guid, currentdocNum, "112", "Purchasing", "Goods Return draft is created from HHD", "Purchasing A/P - Goods Return draft is created from HHD");
                return result;
            }
            else
            {
                return errMsg;
            }
        }

        [WebMethod]
        public DataSet GetGRPOItem(string docNum)
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select LineNum, ItemCode, Dscription 'ItemName',null 'Location',WhsCode,null'BatchNo',OpenQty 'Quantity',null'ManuDate',null'BatchQty' from PDN1 where OpenQty>0 AND DocEntry=(select DocEntry from OPDN where DocStatus='O' AND DocNum=" + docNum + ")";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }

        [WebMethod]
        public DataSet GetGRPOBatchItem(string docEntry)
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.StoredProcedure;
            this.sqlcmd.CommandText = "BatchTransGRPOQty";
            this.sqlcmd.Parameters.AddWithValue("@DocKey", docEntry);
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }
       
        #endregion

        #region---- Iventory Goods Receipt ----
        
        /// <summary>
        /// This Method is used to create the Goods Receipt Draft
        /// </summary>
        /// <param name="guid">Globally Unique Identifier</param>
        /// <param name="xmlInner"></param>
        /// <returns></returns>
        [WebMethod(EnableSession = true)]
        public string GoodsReceiptDraftCreate(string guid, string xmlInner)
        {
            int retValue = 1;
            string errMsg = "";
            SAPbobsCOM.Documents objInventoryGenEntry;
            objInventoryGenEntry = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            objInventoryGenEntry.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryGenEntry;
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(xmlInner);

            #region --- Assign Header Properties ---

            objInventoryGenEntry.DocDate = DateTime.Now;
            objInventoryGenEntry.DocDueDate = DateTime.Now;
            objInventoryGenEntry.UserFields.Fields.Item("U_DocTrans").Value = "HHD";

            #endregion

            XmlNodeList xmlNodeList = xmlDocument.SelectNodes("GoodsReceipt/LineNum");
            try
            {
                #region --- Assign Lines Properties ---

                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    objInventoryGenEntry.Lines.ItemCode = xmlNode.SelectSingleNode("ItemCode").InnerText;
                    objInventoryGenEntry.Lines.WarehouseCode = xmlNode.SelectSingleNode("WhsCode").InnerText;
                    objInventoryGenEntry.Lines.Quantity = Convert.ToDouble(xmlNode.SelectSingleNode("Quantity").InnerText);
                    objInventoryGenEntry.Lines.UserFields.Fields.Item("U_TempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("Location").InnerText.Trim());
                    if (!string.IsNullOrEmpty(xmlNode.SelectSingleNode("BatchNo").InnerText))
                    {
                        string batchArray = xmlNode.SelectSingleNode("BatchNo").InnerText.TrimEnd(',');
                        string[] batchList = batchArray.Split(',');
                        for (int batch = 0; batch < batchList.Length; batch++)
                        {
                            char[] spliter = new char[] { '(', ')' };
                            string[] batchQty = batchList[batch].Split(spliter);
                            string[] manuDate = xmlNode.SelectSingleNode("ManuDate").InnerText.Split(',');
                            objInventoryGenEntry.Lines.BatchNumbers.BatchNumber = batchQty[0];
                            objInventoryGenEntry.Lines.BatchNumbers.ManufacturingDate = Convert.ToDateTime(manuDate[batch]);
                            objInventoryGenEntry.Lines.BatchNumbers.Quantity = double.Parse(batchQty[1]);
                            objInventoryGenEntry.Lines.BatchNumbers.Add();
                        }
                    }
                    objInventoryGenEntry.Lines.Add();
                }

                #endregion

                retValue = objInventoryGenEntry.Add();
                ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetLastError(out retValue, out errMsg);
                if (retValue == -10)
                {
                    errMsg = "Items Have No Stock in the Warehouse";
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objInventoryGenEntry);
            }
            if (retValue == 0)
            {
                string currentdocNum = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey();
                string result = TriggerAlertMessage(guid, currentdocNum, "112", "Inventory", "Goods Receipt draft is created from HHD", "Inventory - Goods Receipt draft is created from HHD");
                return result;
            }
            else
            {
                return errMsg;
            }
        }
        
        #endregion

        #region---- Inventory Goods Issue ----
        
        /// <summary>
        /// This Method is used to create the Goods Issue Draft
        /// </summary>
        /// <param name="guid">Globally Unique Identifier</param>
        /// <param name="xmlInner"></param>
        /// <returns></returns>

        [WebMethod(EnableSession = true)]
        public string GoodsIssueDraftCreate(string guid, string xmlInner)
        {
            int retValue = 1;
            string errMsg = "";
            SAPbobsCOM.Documents objInventoryGenExit;
            objInventoryGenExit = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            objInventoryGenExit.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryGenExit;

            #region --- Assign Header Properties ---

            objInventoryGenExit.UserFields.Fields.Item("U_DocTrans").Value = "HHD";

            #endregion

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(xmlInner);

            XmlNodeList xmlNodeList = xmlDocument.SelectNodes("GoodsIssue/LineNum");
            try
            {
                #region --- Assign Lines Properties ---

                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    objInventoryGenExit.Lines.ItemCode = xmlNode.SelectSingleNode("ItemCode").InnerText;
                    objInventoryGenExit.Lines.WarehouseCode = xmlNode.SelectSingleNode("WhsCode").InnerText;
                    objInventoryGenExit.Lines.Quantity = Convert.ToDouble(xmlNode.SelectSingleNode("Quantity").InnerText);
                    objInventoryGenExit.Lines.UserFields.Fields.Item("U_TempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("Location").InnerText.Trim());
                    if (!string.IsNullOrEmpty(xmlNode.SelectSingleNode("BatchNo").InnerText))
                    {
                        string batchArray = xmlNode.SelectSingleNode("BatchNo").InnerText.TrimEnd(',');
                        string[] batchList = batchArray.Split(',');
                        for (int batch = 0; batch < batchList.Length; batch++)
                        {
                            char[] spliter = new char[] { '(', ')' };
                            string[] batchQty = batchList[batch].Split(spliter);
                            objInventoryGenExit.Lines.BatchNumbers.BatchNumber = batchQty[0];
                            objInventoryGenExit.Lines.BatchNumbers.Quantity = double.Parse(batchQty[1]);
                            objInventoryGenExit.Lines.BatchNumbers.Add();
                        }
                    }
                    objInventoryGenExit.Lines.Add();
                }
                #endregion

                retValue = objInventoryGenExit.Add();
                ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetLastError(out retValue, out errMsg);
                if (retValue == -10)
                {
                    errMsg = "Items Have No Stock in the Warehouse";
                }

            }
            catch (Exception ex)
            {
                errMsg = ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objInventoryGenExit);
            }
            if (retValue == 0)
            {
                string currentdocNum = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey();
                string result = TriggerAlertMessage(guid, currentdocNum, "112", "Inventory", "Goods Issue draft is created from HHD", "Inventory - Goods Issue draft is created from HHD");
                return result;
            }
            else
            {
                return errMsg;
            }
        }
        
        #endregion

        #region ---- Inventory Transfer ----

        /// <summary>
        /// This Method is used to create the Inventory Transfer Draft
        /// </summary>
        /// <param name="guid">Globally Unique Identifier</param>
        /// <param name="fromWhsCode"></param>
        /// <param name="xmlInner"></param>
        /// <returns></returns>
        /// 
        [WebMethod(EnableSession = true)]
        public string InventoryTransferDraftCreate(string guid, string fromWhsCode, string xmlInner)
        {
            string errMsg = "";
            int retValue = 1;
            SAPbobsCOM.StockTransfer objStockTransfer = null;
            try
            {
                
                objStockTransfer = (SAPbobsCOM.StockTransfer)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransferDraft);
                objStockTransfer.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;

                #region --- Assign Header Properties ---

                objStockTransfer.FromWarehouse = fromWhsCode;
                objStockTransfer.UserFields.Fields.Item("U_DocTrans").Value = "HHD";

                #endregion

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlInner);
                XmlNodeList xmlNodeList = xmlDocument.SelectNodes("InventoryTransfer/LineNum");

                #region --- Assign Lines Properties ---

                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    objStockTransfer.Lines.ItemCode = xmlNode.SelectSingleNode("ItemCode").InnerText;
                    objStockTransfer.Lines.WarehouseCode = xmlNode.SelectSingleNode("WhsCode").InnerText;
                    objStockTransfer.Lines.Quantity = Convert.ToDouble(xmlNode.SelectSingleNode("Quantity").InnerText.ToString());
                    objStockTransfer.Lines.UserFields.Fields.Item("U_TempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("Location").InnerText.Trim());
                    objStockTransfer.Lines.UserFields.Fields.Item("U_ToTempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("ToLocation").InnerText.Trim());
                    if (!string.IsNullOrEmpty(xmlNode.SelectSingleNode("BatchNo").InnerText))
                    {
                        string batchArray = xmlNode.SelectSingleNode("BatchNo").InnerText.TrimEnd(',');
                        string[] batchList = batchArray.Split(',');
                        for (int batch = 0; batch < batchList.Length; batch++)
                        {
                            char[] spliter = new char[] { '(', ')' };
                            string[] batchQty = batchList[batch].Split(spliter);
                            objStockTransfer.Lines.BatchNumbers.BatchNumber = batchQty[0];
                            objStockTransfer.Lines.BatchNumbers.Quantity = double.Parse(batchQty[1]);
                            objStockTransfer.Lines.BatchNumbers.Add();
                        }
                    }
                    objStockTransfer.Lines.Add();
                }

                #endregion

                retValue = objStockTransfer.Add();
                ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetLastError(out retValue, out errMsg);
                if (retValue == -10)
                {
                    errMsg = "Items Have No Stock in the Warehouse";
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objStockTransfer);
            }
            if (retValue == 0)
            {
                string currentdocNum = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey();
                string result = TriggerAlertMessage(guid, currentdocNum, "112", "Inventory", "Inventory Transfer draft is created from HHD", "Inventory - Inventory Transfer draft is created from HHD");
                return result;
            }
            else
            {
                return errMsg;
            }
        }

        #endregion

        #region ---- A/R Delivery ----
       
        [WebMethod]
        public DataSet GetSOItem(string docNum)
        {
            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select LineNum,ItemCode, Dscription 'ItemName',null 'Location',WhsCode,null'BatchNo',OpenQty 'Quantity',null'ManuDate',null'BatchQty' from RDR1 where OpenQty > 0 AND DocEntry=(select DocEntry from ORDR where DocStatus='O' AND DocNum=" + docNum + ")";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }

        [WebMethod]
        public DataSet GetSOBatchItem(string docEntry)
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.StoredProcedure;
            this.sqlcmd.CommandText = "BatchTransSOQty";
            this.sqlcmd.Parameters.AddWithValue("@DocKey", docEntry);
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }
       
        /// <summary>
        /// This Method is used to create the A/R Delivery Draft
        /// </summary>
        /// <param name="guid">Globally Unique Identifier</param>
        /// <param name="docEntry"></param>
        /// <param name="xmlInner"></param>
        /// <returns></returns>

        [WebMethod(EnableSession = true)]
        public string ARDeliveryDraftCreate(string guid, int docEntry, string xmlInner)
        {
            string errMsg = "";
            int retValue = 1;
            SAPbobsCOM.Documents objOrder = null;
            SAPbobsCOM.Documents objDeliveryNote = null;
            string lineNumber = "";
            try
            {

                objOrder = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlInner);
                XmlNodeList xmlNodeList = xmlDocument.SelectNodes("ARDelivery/LineNum");
                objOrder.GetByKey(docEntry);
                objDeliveryNote = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                objDeliveryNote.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDeliveryNotes;

                #region --- Assign Header Properties ---

                objDeliveryNote.CardCode = objOrder.CardCode;
                objDeliveryNote.CardName = objOrder.CardName;
                objDeliveryNote.DocDate = DateTime.Now;
                objDeliveryNote.DocDueDate = DateTime.Now;
                objDeliveryNote.TaxDate = DateTime.Now;
                objDeliveryNote.Comments = "Based On Sales Orders " + objOrder.DocNum + ".";
                objDeliveryNote.SalesPersonCode = objOrder.SalesPersonCode;
                objDeliveryNote.UserFields.Fields.Item("U_DocTrans").Value = "HHD";

                #endregion

                #region --- Assign Lines Properties ---

                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    XmlAttribute lID = xmlNode.Attributes["LineNum"];
                    if (lID != null)
                    {
                        lineNumber = lID.Value;
                    }
                    objDeliveryNote.Lines.ItemCode = xmlNode.SelectSingleNode("ItemCode").InnerText;
                    objDeliveryNote.Lines.WarehouseCode = xmlNode.SelectSingleNode("WhsCode").InnerText;
                    objDeliveryNote.Lines.BaseType = 17;
                    objDeliveryNote.Lines.BaseEntry = objOrder.DocEntry;
                    objDeliveryNote.Lines.BaseLine = (Convert.ToInt32(lineNumber));
                    objDeliveryNote.Lines.AccountCode = objOrder.Lines.AccountCode;
                    objDeliveryNote.Lines.Quantity = Convert.ToDouble(xmlNode.SelectSingleNode("Quantity").InnerText);
                    objDeliveryNote.Lines.UserFields.Fields.Item("U_TempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("Location").InnerText.Trim());
                    if (!string.IsNullOrEmpty(xmlNode.SelectSingleNode("BatchNo").InnerText))
                    {
                        string batchArray = xmlNode.SelectSingleNode("BatchNo").InnerText.TrimEnd(',');
                        string[] batchList = batchArray.Split(',');
                        for (int batch = 0; batch < batchList.Length; batch++)
                        {
                            char[] spliter = new char[] { '(', ')' };
                            string[] batchQty = batchList[batch].Split(spliter);
                            objDeliveryNote.Lines.BatchNumbers.BatchNumber = batchQty[0];
                            objDeliveryNote.Lines.BatchNumbers.Quantity = double.Parse(batchQty[1]);
                            objDeliveryNote.Lines.BatchNumbers.Add();
                        }
                    }
                    objDeliveryNote.Lines.Add();
                }

                #endregion

                retValue = objDeliveryNote.Add();
                ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetLastError(out retValue, out errMsg);
                if (retValue == -10)
                {
                    errMsg = "Items Have No Stock in the Warehouse";
                }
                if (retValue == 0)
                {

                }
            }
            catch (Exception ex)
            {
                errMsg = ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objDeliveryNote);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objOrder);
            }

            if (retValue == 0)
            {
                string currentdocNum = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey();
                string result = TriggerAlertMessage(guid, currentdocNum, "112", "Sales", "Delivery draft is created from HHD", "Sales A/R - Delivery draft is created from HHD");
                return result;
            }
            else
            {
                return errMsg;
            }
        }

        #endregion

        #region---- A/R Return ----
       
        /// <summary>
        /// This Method is used to create the A/R Return Draft
        /// </summary>
        /// <param name="guid">Globally Unique Identifier</param>
        /// <param name="docEntry"></param>
        /// <param name="xmlInner"></param>
        /// <returns></returns>

        [WebMethod(EnableSession = true)]
        public string ARGoodsReturnDraftCreate(string guid, int docEntry, string xmlInner)
        {
            string errMsg = "";
            int retValue = 1;
            SAPbobsCOM.Documents objDeliveryNote = null;
            SAPbobsCOM.Documents objReturn = null;
            string lineNumber = "";
            try
            {

                objDeliveryNote = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlInner);
                XmlNodeList xmlNodeList = xmlDocument.SelectNodes("ARGoodsreturn/LineNum");
                objDeliveryNote.GetByKey(docEntry);
                objReturn = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                objReturn.DocObjectCode = SAPbobsCOM.BoObjectTypes.oReturns;

                #region --- Assign Header Properties ---

                objReturn.CardCode = objDeliveryNote.CardCode;
                objReturn.CardName = objDeliveryNote.CardName;
                objReturn.DocDate = DateTime.Now;
                objReturn.DocDueDate = DateTime.Now;
                objReturn.TaxDate = DateTime.Now;
                objReturn.Comments = objDeliveryNote.Comments + ". Based On Deliveries " + objDeliveryNote.DocNum + ".";
                objReturn.UserFields.Fields.Item("U_DocTrans").Value = "HHD";

                #endregion

                #region --- Assign Lines Properties ---

                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    XmlAttribute lID = xmlNode.Attributes["LineNum"];
                    if (lID != null)
                    {
                        lineNumber = lID.Value;
                    }
                    objDeliveryNote.Lines.SetCurrentLine((Convert.ToInt32(lineNumber)));
                    objReturn.Lines.ItemCode = xmlNode.SelectSingleNode("ItemCode").InnerText;
                    objReturn.Lines.ItemDescription = objDeliveryNote.Lines.ItemDescription;
                    objReturn.Lines.WarehouseCode = xmlNode.SelectSingleNode("WhsCode").InnerText;
                    objReturn.Lines.BaseType = 15;
                    objReturn.Lines.BaseEntry = objDeliveryNote.DocEntry;
                    objReturn.Lines.BaseLine = objDeliveryNote.Lines.LineNum;
                    objReturn.Lines.AccountCode = objDeliveryNote.Lines.AccountCode;
                    objReturn.Lines.Quantity = Convert.ToDouble(xmlNode.SelectSingleNode("Quantity").InnerText);
                    objReturn.Lines.UserFields.Fields.Item("U_TempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("Location").InnerText.Trim());
                    if (!string.IsNullOrEmpty(xmlNode.SelectSingleNode("BatchNo").InnerText))
                    {
                        string batchArray = xmlNode.SelectSingleNode("BatchNo").InnerText.TrimEnd(',');
                        string[] batchList = batchArray.Split(',');
                        for (int batch = 0; batch < batchList.Length; batch++)
                        {
                            char[] spliter = new char[] { '(', ')' };
                            string[] batchQty = batchList[batch].Split(spliter);
                            objReturn.Lines.BatchNumbers.BatchNumber = batchQty[0];
                            objReturn.Lines.BatchNumbers.Quantity = double.Parse(batchQty[1]);
                            objReturn.Lines.BatchNumbers.Add();
                        }
                    }

                    objReturn.Lines.Add();
                }

                #endregion

                retValue = objReturn.Add();
                ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetLastError(out retValue, out errMsg);
                if (retValue == -10)
                {
                    errMsg = "Items Have No Stock in the Warehouse";
                }
                if (retValue == 0)
                {
                    objDeliveryNote.Remove();
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objReturn);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objDeliveryNote);
            }
            if (retValue == 0)
            {
                string currentdocNum = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey();
                string result = TriggerAlertMessage(guid, currentdocNum, "112", "Sales", "Return draft is created from HHD", "Sales A/R - Return draft is created from HHD");
                return result;
            }
            else
            {
                return errMsg;
            }
        }

        [WebMethod]
        public DataSet GetARDeliveryItem(string docNum)
        {
            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select LineNum,ItemCode, Dscription 'ItemName',null 'Location',WhsCode,null'BatchNo',OpenQty 'Quantity',null'ManuDate',null'BatchQty' from DLN1 where OpenQty >0 AND  DocEntry=(select DocEntry from ODLN where DocNum=" + docNum + ")";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                CloseConnection();

            }
            return ds;
        }

        [WebMethod]
        public DataSet GetARDeliveryBatchItem(string docEntry, string docNum)
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select distinct T0.ItemCode,T0.Dscription 'ItemName',null 'Location',T2.WhsCode,T2.BatchNum 'BatchNo',(T2.Quantity- CASE(Select COUNT(ItemCode) from IBT1 where BsDocEntry=T1.DocEntry and BaseType='16' and BatchNum=T2.BatchNum) WHEN 0 THEN 0 ELSE (Select Quantity from IBT1 where BsDocEntry=T1.DocEntry and BaseType='16' and BatchNum=T2.BatchNum) END) AS Quantity,null'ManuDate',null'BatchQty' from DLN1 as T0 inner join ODLN as T1 on T1.DocEntry = T0.DocEntry inner Join IBT1 as T2 on T2.BaseEntry=T1.DocEntry and T2.ItemCode=t0.ItemCode where t2.Direction=1 and T1.DocNum=" + docNum + " and  T1.DocEntry=" + docEntry + "";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                CloseConnection();

            }
            return ds;
        }
        
        #endregion
        
        #region---- A/R CreditMemo ----
        
        /// <summary>
        /// This Method is used to create the A/R CreditMemo Draft
        /// </summary>
        /// <param name="guid">Globally Unique Identifier</param>
        /// <param name="docEntry"></param>
        /// <param name="xmlInner"></param>
        /// <returns></returns>

        [WebMethod(EnableSession = true)]
        public string ARCreditMemoDraftCreate(string guid, int docEntry, string xmlInner)
        {
            string errMsg = "";
            int retValue = 1;
            string lineNumber = "";
            SAPbobsCOM.Documents objInvoice = null;
            SAPbobsCOM.Documents objCreditNote = null;
            try
            {
                objInvoice = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlInner);
                XmlNodeList xmlNodeList = xmlDocument.SelectNodes("ARCreditMemo/LineNum");
                objInvoice.GetByKey(docEntry);
                objCreditNote = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                objCreditNote.DocObjectCode = SAPbobsCOM.BoObjectTypes.oCreditNotes;

                #region --- Assign Header Properties ---

                objCreditNote.ContactPersonCode = objInvoice.ContactPersonCode;
                objCreditNote.SummeryType = objInvoice.SummeryType;
                objCreditNote.PayToCode = objInvoice.PayToCode;
                objCreditNote.PayToBankCountry = objInvoice.PayToBankCountry;
                objCreditNote.PayToBankCode = objInvoice.PayToBankCode;
                objCreditNote.PayToBankAccountNo = objInvoice.PayToBankAccountNo;
                objCreditNote.PayToBankBranch = objInvoice.PayToBankBranch;
                objCreditNote.PaymentBlockEntry = objInvoice.PaymentBlockEntry;
                objCreditNote.PaymentMethod = objInvoice.PaymentMethod;
                objCreditNote.DiscountPercent = objInvoice.DiscountPercent;
                objCreditNote.Project = objInvoice.Project;
                objCreditNote.DocType = objInvoice.DocType;
                objCreditNote.RequriedDate = objInvoice.RequriedDate;
                objCreditNote.Address = objInvoice.Address;
                objCreditNote.NumAtCard = objInvoice.NumAtCard;
                objCreditNote.CardCode = objInvoice.CardCode;
                objCreditNote.CardName = objInvoice.CardName;
                objCreditNote.DocDate = DateTime.Now;
                objCreditNote.DocDueDate = DateTime.Now;
                objCreditNote.TaxDate = DateTime.Now;
                objCreditNote.SalesPersonCode = objInvoice.SalesPersonCode;
                objCreditNote.UserFields.Fields.Item("U_DocTrans").Value = "HHD";

                #endregion

                #region --- Assign Lines Properties ---

                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    XmlAttribute lID = xmlNode.Attributes["LineNum"];
                    if (lID != null)
                    {
                        lineNumber = lID.Value;
                    }
                    objInvoice.Lines.SetCurrentLine((Convert.ToInt32(lineNumber)));
                    objCreditNote.Lines.ItemCode = xmlNode.SelectSingleNode("ItemCode").InnerText;
                    objCreditNote.Lines.ItemDescription = xmlNode.SelectSingleNode("ItemName").InnerText;
                    objCreditNote.Lines.Price = objInvoice.Lines.Price;
                    objCreditNote.Lines.PriceAfterVAT = objInvoice.Lines.PriceAfterVAT;
                    objCreditNote.Lines.DiscountPercent = objInvoice.Lines.DiscountPercent;

                    objCreditNote.Lines.WarehouseCode = xmlNode.SelectSingleNode("WhsCode").InnerText;
                    objCreditNote.Lines.Quantity = Convert.ToDouble(xmlNode.SelectSingleNode("Quantity").InnerText);
                    objCreditNote.Lines.UserFields.Fields.Item("U_TempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("Location").InnerText.Trim());
                    if (!string.IsNullOrEmpty(xmlNode.SelectSingleNode("BatchNo").InnerText))
                    {
                        string batchArray = xmlNode.SelectSingleNode("BatchNo").InnerText.TrimEnd(',');
                        string[] batchList = batchArray.Split(',');
                        for (int batch = 0; batch < batchList.Length; batch++)
                        {
                            char[] spliter = new char[] { '(', ')' };
                            string[] batchQty = batchList[batch].Split(spliter);
                            objCreditNote.Lines.BatchNumbers.BatchNumber = batchQty[0];
                            objCreditNote.Lines.BatchNumbers.Quantity = double.Parse(batchQty[1]);
                            objCreditNote.Lines.BatchNumbers.Add();
                        }
                    }
                    objCreditNote.Lines.Add();
                }

                #endregion

                retValue = objCreditNote.Add();
                ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetLastError(out retValue, out errMsg);
                if (retValue == -10)
                {
                    errMsg = "Items Have No Stock in the Warehouse";
                }
                if (retValue == 0)
                {
                    objInvoice.Remove();
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCreditNote);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objInvoice);
            }
            if (retValue == 0)
            {
                string currentdocNum = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey();
                string result = TriggerAlertMessage(guid, currentdocNum, "112", "Sales", "Credit Memo draft is created from HHD", "Sales A/R - Credit Memo draft is created from HHD");
                return result;
            }
            else
            {
                return errMsg;
            }
        }

        [WebMethod]
        public DataSet GetInvoiceItem(string docNum)
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select LineNum,ItemCode, Dscription 'ItemName',null 'Location',WhsCode,null'BatchNo',OpenQty 'Quantity',null'ManuDate',null'BatchQty' from INV1 where OpenQty > 0 AND DocEntry=(select DocEntry from OINV where DocStatus='O' AND DocNum=" + docNum + ")";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }

        [WebMethod]
        public DataSet GetInvoiceBaseItem(string docNum)
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "SELECT DocEntry,BaseRef, BaseEntry,BaseType FROM INV1 WHERE BaseType=15 and DocEntry=(SELECT DocEntry FROM OINV WHERE DocNum=" + docNum + ")";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }

        #endregion

        #region---- Receipt from Production ----

        /// <summary>
        /// This Method is used to create the Receipt from Production
        /// </summary>
        /// <param name="guid">Globally Unique Identifier</param>
        /// <param name="xmlInner"></param>
        /// <returns></returns>
        [WebMethod(EnableSession = true)]
        public string ProductionReceiptCreate(string guid, int docEntry, string xmlInner)
        {
            string errMsg = "";
            int retValue = 1;
            SAPbobsCOM.ProductionOrders objProductionOrder = null;
            SAPbobsCOM.Documents objInventoryGenEntry = null;
            try
            {
                objProductionOrder = (SAPbobsCOM.ProductionOrders)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlInner);
                XmlNodeList xmlNodeList = xmlDocument.SelectNodes("ProductionReceipt/LineNum");
                objProductionOrder.GetByKey(docEntry);
                objInventoryGenEntry = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

                #region --- Assign Header Properties ---

                objInventoryGenEntry.UserFields.Fields.Item("U_DocTrans").Value = "HHD";

                #endregion

                #region --- Assign Lines Properties ---

                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    objInventoryGenEntry.Lines.WarehouseCode = xmlNode.SelectSingleNode("WhsCode").InnerText;
                    objInventoryGenEntry.Lines.BaseType = 202;
                    objInventoryGenEntry.Lines.BaseEntry = objProductionOrder.Lines.DocumentAbsoluteEntry;
                    objInventoryGenEntry.Lines.Quantity = Convert.ToDouble(xmlNode.SelectSingleNode("Quantity").InnerText);
                    objInventoryGenEntry.Lines.UserFields.Fields.Item("U_TempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("Location").InnerText.Trim());
                    if (!string.IsNullOrEmpty(xmlNode.SelectSingleNode("BatchNo").InnerText))
                    {
                        string batchArray = xmlNode.SelectSingleNode("BatchNo").InnerText.TrimEnd(',');
                        string[] batchList = batchArray.Split(',');
                        for (int batch = 0; batch < batchList.Length; batch++)
                        {
                            char[] spliter = new char[] { '(', ')' };
                            string[] batchQty = batchList[batch].Split(spliter);
                            string[] manuDate = xmlNode.SelectSingleNode("ManuDate").InnerText.Split(',');
                            objInventoryGenEntry.Lines.BatchNumbers.BatchNumber = batchQty[0];
                            objInventoryGenEntry.Lines.BatchNumbers.ManufacturingDate = Convert.ToDateTime(manuDate[batch]);
                            objInventoryGenEntry.Lines.BatchNumbers.Quantity = double.Parse(batchQty[1]);
                            objInventoryGenEntry.Lines.BatchNumbers.Add();
                        }
                    }
                    objInventoryGenEntry.Lines.Add();
                }

                #endregion

                retValue = objInventoryGenEntry.Add();
                ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetLastError(out retValue, out errMsg);
                if (retValue == -10)
                {
                    errMsg = "Items Have No Stock in the Warehouse";
                }
                if (retValue == 0)
                {
                    string DocEntry = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey().ToString();
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.ToString();
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(objInventoryGenEntry);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objProductionOrder);
            }
            if (retValue == 0)
            {
                string currentdocNum = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey();
                string result = TriggerAlertMessage(guid, currentdocNum, "59", "Production", "Receipt from Production is created from HHD", "Production - Receipt from Production is created from HHD");
                return result;
            }
            else
            {
                return errMsg;
            }
        }

        [WebMethod]
        public DataSet GetProductionOrderforReceiptItem(string docNum)
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select ItemCode, (Select ItemName from OITM T1 Where T1.ItemCode=T0.ItemCode) 'ItemName',null 'Location',T0.wareHouse 'WhsCode',null'BatchNo',(PlannedQty-CmpltQty) 'Quantity',null'ManuDate',null'BatchQty' from OWOR T0 where [Status]='R' AND DocNum=" + docNum + "";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }

        [WebMethod]
        public DataSet GetProductionOrderStatus(string docNum)
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select [Status] from OWOR where DocNum=" + docNum + "";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }
        #endregion

        #region---- Issue for Production ----

        /// <summary>
        /// This Method is used to create the Issue for Production
        /// </summary>
        /// <param name="guid">Globally Unique Identifier</param>
        /// <param name="xmlInner"></param>
        /// <returns></returns>

        [WebMethod(EnableSession = true)]
        public string ProductionIssueCreate(string guid, int docEntry, string xmlInner)
        {
            string errMsg = "";
            int lineNumber = 0;
            int retValue = 1;
            SAPbobsCOM.ProductionOrders objProductionOrder = null;
            SAPbobsCOM.Documents objInventoryGenExit = null;
            try
            {
                objProductionOrder = (SAPbobsCOM.ProductionOrders)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlInner);
                XmlNodeList xmlNodeList = xmlDocument.SelectNodes("ProductionIssue/LineNum");
                objProductionOrder.GetByKey(docEntry);
                objInventoryGenExit = (SAPbobsCOM.Documents)((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

                #region --- Assign Header Properties ---

                objInventoryGenExit.UserFields.Fields.Item("U_DocTrans").Value = "HHD";

                #endregion

                #region --- Assign Lines Properties ---

                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    XmlAttribute lineNum = xmlNode.Attributes["LineNum"];
                    if (lineNum != null)
                    {
                        lineNumber = int.Parse(lineNum.Value);
                    }
                    objInventoryGenExit.Lines.WarehouseCode = xmlNode.SelectSingleNode("WhsCode").InnerText;
                    objInventoryGenExit.Lines.BaseType = 202;
                    objInventoryGenExit.Lines.BaseEntry = objProductionOrder.Lines.DocumentAbsoluteEntry;
                    objInventoryGenExit.Lines.BaseLine = lineNumber;
                    objInventoryGenExit.Lines.Quantity = Convert.ToDouble(xmlNode.SelectSingleNode("Quantity").InnerText);
                    objInventoryGenExit.Lines.UserFields.Fields.Item("U_TempLocation").Value = RemoveColon(xmlNode.SelectSingleNode("Location").InnerText.Trim());
                    if (!string.IsNullOrEmpty(xmlNode.SelectSingleNode("BatchNo").InnerText))
                    {
                        string batchArray = xmlNode.SelectSingleNode("BatchNo").InnerText.TrimEnd(',');
                        string[] batchList = batchArray.Split(',');
                        for (int batch = 0; batch < batchList.Length; batch++)
                        {
                            char[] spliter = new char[] { '(', ')' };
                            string[] batchQty = batchList[batch].Split(spliter);
                            objInventoryGenExit.Lines.BatchNumbers.BatchNumber = batchQty[0];
                            objInventoryGenExit.Lines.BatchNumbers.Quantity = double.Parse(batchQty[1]);
                            objInventoryGenExit.Lines.BatchNumbers.Add();
                        }
                    }
                    objInventoryGenExit.Lines.Add();
                }

                #endregion

                retValue = objInventoryGenExit.Add();
                ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetLastError(out retValue, out errMsg);
                if (retValue == -10)
                {
                    errMsg = "Items Have No Stock in the Warehouse";
                }
                if (retValue == 0)
                {

                }
            }
            catch (Exception ex)
            {
                errMsg = ex.ToString();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objInventoryGenExit);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objProductionOrder);
            }
            if (retValue == 0)
            {
                string currentdocNum = ((SAPbobsCOM.Company)HttpContext.Current.Cache[guid]).GetNewObjectKey();
                string result = TriggerAlertMessage(guid, currentdocNum, "60", "Production", "Issue for Production is created from HHD", "Production - Issue for Production is created from HHD");
                return result;
            }
            else
            {
                return errMsg;
            }
        }

        [WebMethod]
        public DataSet GetProductionOrderforIssueItem(string docNum)
        {

            DataSet ds = new DataSet();
            GetConnection();
            this.sqlcmd.CommandType = CommandType.Text;
            this.sqlcmd.CommandText = "select T0.LineNum,ItemCode, (Select ItemName from OITM T1 Where T1.ItemCode=T0.ItemCode) 'ItemName',null 'Location',T0.wareHouse 'WhsCode',null'BatchNo',(PlannedQty - IssuedQty) 'Quantity',null'ManuDate',null'BatchQty' from WOR1 T0 where DocEntry=(select DocEntry from OWOR where [Status]='R' AND DocNum=" + docNum + ")";
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
                da.Fill(ds);
            }
            catch (SqlException SqlCee)
            {
                throw SqlCee;

            }
            finally
            {
                this.sqlcmd.Parameters.Clear();
                CloseConnection();

            }
            return ds;
        }
        #endregion
    }

}



