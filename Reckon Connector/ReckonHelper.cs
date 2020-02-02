using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Newtonsoft.Json.Converters;
using QBXMLRP2Lib;
using ReckonDesktop.Model;
using Reckon_Connector;

namespace Reckon_Connector
{
    public class ReckonHelper
    {
        public string appID = "CWRecAccDesk";
        public string appName = "Colourworks Reckon Accounts Desktop Connector";
        public string ticket;
        public RequestProcessor2 rp = new RequestProcessor2();
        //public string maxVersion = "7.1";
        public string maxVersion = "6.1";
        public QBFileMode mode = QBFileMode.qbFileOpenDoNotCare;


        //THIS IS INVALUABLE
        //https://developer-static.intuit.com/qbSDK-current/Common/newOSR/index.html


        /*ALSO THIS NEEDS TO BE LOOKED AT AT SOME POINT - This is how we will map autofile records with Reckon invoices and bills for updating and adding new items to exisintg bills
        
        defMacro
        Note that defMacro was introduced with SDK 2.0, so it only works with 2.0 and higher requests. You can use the defMacro attribute to assign a name to the TxnID or TxnListID that this aggregate will return. This way you can refer to the transaction by name in a later request. For example, if you were using the qbXML API and you defined an invoice add request with the name TxnID:RecvPmt1234, as shown below, then you could refer to that invoice by name in a later receive payment add request:

        <InvoiceAddRq>
        <InvoiceAdd defMacro= "TxnID:RecvPmt1234">
        . . .
        <ReceivePaymentAddRq>
        <ReceivePaymentAdd>
        . . .
        <TxnID useMacro="TxnID:RecvPmt1234"/>
        . . .

        If you use macros with QBOE...
        There may be a buug in the QBOE implementation of this feature. If this feature is not working for you in QBOE, try stripping the prefix "TxnID:" from the name of the useMacro. For example, defMacro="TxnID:RecvPmt1234" and useMacro="RecvPmt1234"
        */



        public bool connectToQB(string companyFile)
        {


            //int authflags = 0;
            //authflags = authflags | 0x01;



            //int authFlags = 0x1 - 0x2 - 0x4;

            //var qbXMLCOM = new RequestProcessor2();
            //var prefs = new AuthPreferences();
            //prefs = qbXMLCOM.AuthPreferences as AuthPreferences;
            //prefs.PutAuthFlags(authFlags);

            //qbXMLCOM.OpenConnection(appID, appName);

            //ticket = qbXMLCOM.BeginSession(companyFile, mode);
            //var versions = qbXMLCOM.get_QBXMLVersionsForSession(ticket);
            //maxVersion = versions.GetValue(versions.Length - 1).ToString();
            //return true;


            //Dim authFlags As Long
            //authFlags = 0
            //authFlags = authFlags Or & H8 &
            //authFlags = authFlags Or & H4 &
            //authFlags = authFlags Or & H2 &
            //authFlags = authFlags Or & H1 &
            //authFlags = authFlags Or & H80000000
            //Dim qbXMLCOM As QBXMLRP2Lib.RequestProcessor2
            //Dim prefs As QBXMLRP2Lib.AuthPreferences
            //Set prefs = qbXMLCOM.AuthPreferences
            //prefs.PutAuthFlags(authFlags)
            int authFlags = 0;
            authFlags = 0;
            authFlags = (int) (authFlags | 0x8L);
            authFlags = (int) (authFlags | 0x4L);
            authFlags = (int) (authFlags | 0x2L);
            authFlags = (int) (authFlags | 0x1L);
//            int authFlags = 0x1 - 0x2 - 0x4;
            //var rp = new RequestProcessor2();
            var prefs = (AuthPreferences) rp.AuthPreferences; //new AuthPreferences();
//            prefs = (AuthPreferences)qbXMLCOM.AuthPreferences;
            prefs.PutAuthFlags(authFlags);
            try
            {
                rp.OpenConnection(appID, appName);
                ticket = rp.BeginSession(companyFile, mode);
                var versions = rp.get_QBXMLVersionsForSession(ticket);
                maxVersion = versions.GetValue(versions.Length - 1).ToString();
            }
            catch (Exception Ex)
            {
                if (Ex.Message != null && Ex.Message == "Could not start Reckon Accounts.")
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine + "Please ensure you have Reckon running and the company file you wish to integrate open.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }
                rp.CloseConnection();
                return false;
            }
            return true;
        }
      
        public bool disconnectFromQB()
        {
            if (ticket != null)
            {
                try
                {
                    rp.EndSession(ticket);
                    ticket = null;
                    rp.CloseConnection();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
            return true;
        }



        
        public string getDateString(DateTime dt)
        {
            string year = dt.Year.ToString();
            string month = dt.Month.ToString();
            if (month.Length < 2) month = "0" + month;
            string day = dt.Day.ToString();
            if (day.Length < 2) day = "0" + day;
            return year + "-" + month + "-" + day;
        }

        public string buildAccountQueryRqXML(string[] includeRetElement, string fullName)
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion ?? "6.1");
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement CustomerQueryRq = xmlDoc.CreateElement("AccountQueryRq");
            qbXMLMsgsRq.AppendChild(CustomerQueryRq);
            if (fullName != null)
            {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                CustomerQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                CustomerQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }
            CustomerQueryRq.SetAttribute("requestID", "7");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        public string buildCustomerQueryRqXML(string[] includeRetElement, string fullName)
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion ?? "6.1");
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement CustomerQueryRq = xmlDoc.CreateElement("CustomerQueryRq");
            qbXMLMsgsRq.AppendChild(CustomerQueryRq);
            if (fullName != null)
            {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                CustomerQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }

            //Uncomment these to show custom fields
            //XmlElement ownerIDRet = xmlDoc.CreateElement("OwnerID");
            //CustomerQueryRq.AppendChild(ownerIDRet).InnerText = "0";
            
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                CustomerQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }

            CustomerQueryRq.SetAttribute("requestID", "1");
            xml = xmlDoc.OuterXml;
            return xml;
        }
        
        public string buildVendorQueryRqXML(string[] includeRetElement, string fullName)
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion ?? "6.1");
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement CustomerQueryRq = xmlDoc.CreateElement("VendorQueryRq");
            qbXMLMsgsRq.AppendChild(CustomerQueryRq);
            if (fullName != null)
            {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                CustomerQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }

            //Uncomment these to show custom fields
            //XmlElement ownerIDRet = xmlDoc.CreateElement("OwnerID");
            //CustomerQueryRq.AppendChild(ownerIDRet).InnerText = "0";

            for (int x = 0; x < includeRetElement.Length; x++)
            {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                CustomerQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }

            CustomerQueryRq.SetAttribute("requestID", "1");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        public string buildItemQueryRqXML(string[] includeRetElement, string fullName)
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion ?? "6.1");
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement ItemQueryRq = xmlDoc.CreateElement("ItemQueryRq");
            qbXMLMsgsRq.AppendChild(ItemQueryRq);
            if (fullName != null)
            {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                ItemQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                ItemQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }
            ItemQueryRq.SetAttribute("requestID", "2");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        public string buildClassQueryRqXML(string[] includeRetElement, string fullName)
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion ?? "6.1");
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement ItemQueryRq = xmlDoc.CreateElement("ClassQueryRq");
            qbXMLMsgsRq.AppendChild(ItemQueryRq);
            if (fullName != null)
            {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                ItemQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                ItemQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }
            ItemQueryRq.SetAttribute("requestID", "2");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        public IEnumerable<string> parseItemQueryRs(string xml)
        {
            /*
              <?xml version="1.0" ?> 
            - <QBXML>
            - <QBXMLMsgsRs>
            - <ItemQueryRs requestID="2" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
            - <ItemServiceRet>
  	            <ListID>20000-933272655</ListID> 
  	            <TimeCreated>1999-07-29T11:24:15-08:00</TimeCreated> 
  	            <TimeModified>2007-12-15T11:32:53-08:00</TimeModified> 
  	            <EditSequence>1197747173</EditSequence> 
  	            <Name>Installation</Name> 
  	            <FullName>Installation</FullName> 
  	            <IsActive>true</IsActive> 
  	            <Sublevel>0</Sublevel> 
            - 	<SalesTaxCodeRef>
  		            <ListID>20000-999022286</ListID> 
  		            <FullName>Non</FullName> 
  	            </SalesTaxCodeRef>
            - 	<SalesOrPurchase>
  		            <Desc>Installation labor</Desc> 
  		            <Price>35.00</Price> 
            - 		<AccountRef>
  			            <ListID>190000-933270541</ListID> 
  			            <FullName>Construction Income:Labor Income</FullName> 
  		            </AccountRef>
  	            </SalesOrPurchase>
              </ItemServiceRet>
              </ItemQueryRs>
              </QBXMLMsgsRs>
              </QBXML>
            */

            var retVal = new List<string>();
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null)
            {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more)
            {
                switch (nav.LocalName)
                {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ItemQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ItemServiceRet":
                        //more = nav.MoveToFirstChild();
                        //continue;
                    case "ItemNonInventoryRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ItemOtherChargeRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ItemInventoryRet":
                    case "ItemInventoryAssemblyRet":
                    case "ItemFixedAssetRet":
                    case "ItemSubtotalRet":
                    case "ItemDiscountRet":
                    case "ItemPaymentRet":
                    case "ItemSalesTaxRet":
                    case "ItemSalesTaxGroupRet":
                    case "ItemGroupRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "FullName":
                        retVal.Add(nav.Value.Trim());
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    case "SalesOrPurchase":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Desc":
                    case "Price":
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }

        public IEnumerable<string> parseClassQueryRs(string xml)
        {
            /*
              <?xml version="1.0" ?> 
            - <QBXML>
            - <QBXMLMsgsRs>
            - <ItemQueryRs requestID="2" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
            - <ItemServiceRet>
  	            <ListID>20000-933272655</ListID> 
  	            <TimeCreated>1999-07-29T11:24:15-08:00</TimeCreated> 
  	            <TimeModified>2007-12-15T11:32:53-08:00</TimeModified> 
  	            <EditSequence>1197747173</EditSequence> 
  	            <Name>Installation</Name> 
  	            <FullName>Installation</FullName> 
  	            <IsActive>true</IsActive> 
  	            <Sublevel>0</Sublevel> 
            - 	<SalesTaxCodeRef>
  		            <ListID>20000-999022286</ListID> 
  		            <FullName>Non</FullName> 
  	            </SalesTaxCodeRef>
            - 	<SalesOrPurchase>
  		            <Desc>Installation labor</Desc> 
  		            <Price>35.00</Price> 
            - 		<AccountRef>
  			            <ListID>190000-933270541</ListID> 
  			            <FullName>Construction Income:Labor Income</FullName> 
  		            </AccountRef>
  	            </SalesOrPurchase>
              </ItemServiceRet>
              </ItemQueryRs>
              </QBXMLMsgsRs>
              </QBXML>
            */

            var retVal = new List<string>();
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null)
            {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more)
            {
                switch (nav.LocalName)
                {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ClassQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ClassRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "FullName":
                        retVal.Add(nav.Value.Trim());
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }

        public string[] parseInvoiceAddRs(string xml)
        {
            string[] retVal = new string[3];
            try
            {
                XmlNodeList RsNodeList = null;
                XmlDocument Doc = new XmlDocument();
                Doc.LoadXml(xml);
                RsNodeList = Doc.GetElementsByTagName("InvoiceAddRs");
                XmlAttributeCollection rsAttributes = RsNodeList.Item(0).Attributes;
                XmlNode statusCode = rsAttributes.GetNamedItem("statusCode");
                retVal[0] = Convert.ToString(statusCode.Value);
                XmlNode statusSeverity = rsAttributes.GetNamedItem("statusSeverity");
                retVal[1] = Convert.ToString(statusSeverity.Value);
                XmlNode statusMessage = rsAttributes.GetNamedItem("statusMessage");
                retVal[2] = Convert.ToString(statusMessage.Value);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error encountered when parsing Invoice info returned from QuickBooks: " + e.Message);
                retVal = null;
            }
            return retVal;
        }

        public string[] parseBillAddRs(string xml)
        {
            string[] retVal = new string[3];
            try
            {
                XmlNodeList RsNodeList = null;
                XmlDocument Doc = new XmlDocument();
                Doc.LoadXml(xml);
                RsNodeList = Doc.GetElementsByTagName("BillAddRs");
                XmlAttributeCollection rsAttributes = RsNodeList.Item(0).Attributes;
                XmlNode statusCode = rsAttributes.GetNamedItem("statusCode");
                retVal[0] = Convert.ToString(statusCode.Value);
                XmlNode statusSeverity = rsAttributes.GetNamedItem("statusSeverity");
                retVal[1] = Convert.ToString(statusSeverity.Value);
                XmlNode statusMessage = rsAttributes.GetNamedItem("statusMessage");
                retVal[2] = Convert.ToString(statusMessage.Value);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error encountered when parsing Bill info returned from QuickBooks: " + e.Message);
                retVal = null;
            }
            return retVal;
        }
        
        public XmlElement buildRqEnvelope(XmlDocument doc, string maxVer)
        {
            doc.AppendChild(doc.CreateXmlDeclaration("1.0", null, null));
            doc.AppendChild(doc.CreateProcessingInstruction("qbxml", "version=\"" + maxVer + "\""));
            XmlElement qbXML = doc.CreateElement("QBXML");
            doc.AppendChild(qbXML);
            XmlElement qbXMLMsgsRq = doc.CreateElement("QBXMLMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            return qbXMLMsgsRq;
        }




        public string processRequestFromQB(string request)
        {
            try
            {
                return rp.ProcessRequest(ticket, request);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
        }




        public bool AddInvoiceAR(string companyFile, string CustomerRef, string ARAccountRef, string TxDate, string RefNo, string InvLineRef, string InvLineDesc, string InvLineQuantity, string InvLineAmt, ref string Errors )
        {
            var success = true;
            string requestXML = buildInvoiceAddRqXML(CustomerRef, ARAccountRef, TxDate, RefNo, InvLineRef, InvLineDesc, InvLineQuantity, InvLineAmt);
            if (requestXML == null)
            {
                MessageBox.Show("One of the input is missing. Double-check your entries and then click Save again.", "Error saving invoice");
                success = false;
                return false;
            }
            connectToQB(companyFile);
            string response = processRequestFromQB(requestXML);
            disconnectFromQB();
            string[] status = new string[3];
            if (response != null) status = parseInvoiceAddRs(response);
            string msg = "";

            if (status != null)
            {
                if (status[0] == "0")
                {
                    msg = "Invoice was added successfully!";
               
                }
                else
                {
                    msg = "Could not add invoice " + RefNo;
                    success = false;
                }

                msg = msg + "\n\n";
                msg = msg + "Status Code = " + status[0] + "\n";
                msg = msg + "Status Severity = " + status[1] + "\n";
                msg = msg + "Status Message = " + status[2] + "\n";
                if (!success)
                {
                    Errors = Errors + msg;
                }
            }
            else
            {
                msg = "Could not add invoice " + RefNo;
                success = false;
                if (!success)
                {
                    Errors = Errors + msg;
                }
            }

            //MessageBox.Show(msg);
            return success;
        }

        public string buildInvoiceAddRqXML(string CustomerRef, string ARAccountRef, string TxDate, string RefNo, string InvLineRef, string InvLineDesc, string InvLineQuantity, string InvLineAmt)
        {

      string requestXML = "";
            
            //GET ALL INPUT INTO XML
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion ?? "6.1");
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement InvoiceAddRq = xmlDoc.CreateElement("InvoiceAddRq");
            qbXMLMsgsRq.AppendChild(InvoiceAddRq);
            XmlElement InvoiceAdd = xmlDoc.CreateElement("InvoiceAdd");
            InvoiceAddRq.AppendChild(InvoiceAdd);

            // CustomerRef -> FullName
            if (! string.IsNullOrEmpty(CustomerRef))
            {
                XmlElement Element_CustomerRef = xmlDoc.CreateElement("CustomerRef");
                InvoiceAdd.AppendChild(Element_CustomerRef);
                XmlElement Element_CustomerRef_FullName = xmlDoc.CreateElement("ListID");
                Element_CustomerRef.AppendChild(Element_CustomerRef_FullName).InnerText = CustomerRef;
            }

            // AccountRef -> FullName
            if (!string.IsNullOrEmpty(ARAccountRef))
            {
                XmlElement Element_AccountRef = xmlDoc.CreateElement("ARAccountRef");
                InvoiceAdd.AppendChild(Element_AccountRef);
                XmlElement Element_AccountRef_FullName = xmlDoc.CreateElement("FullName");
                Element_AccountRef.AppendChild(Element_AccountRef_FullName).InnerText = ARAccountRef;
            }

            // TxnDate 
            DateTime DT_TxnDate = System.DateTime.Today;
            if (! string.IsNullOrEmpty(TxDate))
            {
                DT_TxnDate = Convert.ToDateTime(TxDate);
                string TxnDate = getDateString(DT_TxnDate);
                XmlElement Element_TxnDate = xmlDoc.CreateElement("TxnDate");
                InvoiceAdd.AppendChild(Element_TxnDate).InnerText = TxnDate;
            }

            // RefNumber 
            if (!string.IsNullOrEmpty(RefNo))
            {
                XmlElement Element_RefNumber = xmlDoc.CreateElement("RefNumber");
                InvoiceAdd.AppendChild(Element_RefNumber).InnerText = RefNo;
            }

            //Invoice Line (REQUIRED)
            XmlElement Element_InvoiceLineAdd;

            Element_InvoiceLineAdd = xmlDoc.CreateElement("InvoiceLineAdd");
            InvoiceAdd.AppendChild(Element_InvoiceLineAdd);
            
            if (!string.IsNullOrEmpty(InvLineRef))
            {
                XmlElement Element_InvoiceLineAdd_ItemRef = xmlDoc.CreateElement("ItemRef");
                Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_ItemRef);
                XmlElement Element_InvoiceLineAdd_ItemRef_FullName = xmlDoc.CreateElement("FullName");
                Element_InvoiceLineAdd_ItemRef.AppendChild(Element_InvoiceLineAdd_ItemRef_FullName).InnerText = InvLineRef;
            }
            if (!string.IsNullOrEmpty(InvLineDesc))
            {
                XmlElement Element_InvoiceLineAdd_Desc = xmlDoc.CreateElement("Desc");
                Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Desc).InnerText = InvLineDesc;
            }
            if (!string.IsNullOrEmpty(InvLineQuantity))
            {
                XmlElement Element_InvoiceLineAdd_Quantity = xmlDoc.CreateElement("Quantity");
                Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Quantity).InnerText = InvLineQuantity;
            }
            if (!string.IsNullOrEmpty(InvLineAmt))
            {
                XmlElement Element_InvoiceLineAdd_Amount = xmlDoc.CreateElement("Amount");
                Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Amount).InnerText = InvLineAmt;
            }


            InvoiceAddRq.SetAttribute("requestID", "99");
            requestXML = xmlDoc.OuterXml;

            return requestXML;
        }




        public bool AddBillAP(string companyFile, string VendorRef, string APAccountRef, string TxDate, string RefNo, string InvLineRef, string InvLineDesc, string InvLineQuantity, string InvLineAmt, string InvLineCustomer, string InvLineClass, ref string Errors)
        {
            bool success = true;
            string requestXML = buildBillAddRqXML(VendorRef, APAccountRef, TxDate, RefNo, InvLineRef, InvLineDesc, InvLineQuantity, InvLineAmt, InvLineCustomer, InvLineClass);
            if (requestXML == null)
            {
                MessageBox.Show("One of the input is missing. Double-check your entries and then click Save again.", "Error saving invoice");
                success = false;
                return success;
            }
            connectToQB(companyFile);
            string response = processRequestFromQB(requestXML);
            disconnectFromQB();
            string[] status = new string[3];
            if (response != null) status = parseBillAddRs(response);
            string msg = "";

            if (response != null & status[0] == "0")
            {
                msg = "Bill was added successfully!";
            }
            else
            {
                msg = "Could not add Bill " + RefNo;
                success = false;
            }

            msg = msg + "\n\n";
            msg = msg + "Status Code = " + status[0] + "\n";
            msg = msg + "Status Severity = " + status[1] + "\n";
            msg = msg + "Status Message = " + status[2] + "\n";
            //MessageBox.Show(msg);
            if (!success)
            {
                Errors = Errors + msg;
            }
            return success;
        }

        public bool AddBillAPLines(string companyFile, string VendorRef, string APAccountRef, string TxDate, string RefNo, string InvLineRef, string InvLineDesc, string InvLineQuantity, string InvLineAmt, string InvLineCustomer, string InvLineClass, List<AccountingLine> lines, ref string Errors)
        {
            bool success = true;
            string requestXML = buildBillAddRqXML(VendorRef, APAccountRef, TxDate, RefNo, InvLineRef, InvLineDesc, InvLineQuantity, InvLineAmt, InvLineCustomer, InvLineClass, lines);
            if (requestXML == null)
            {
                MessageBox.Show("One of the input is missing. Double-check your entries and then click Save again.", "Error saving invoice");
                success = false;
                return success;
            }
            connectToQB(companyFile);
            string response = processRequestFromQB(requestXML);
            disconnectFromQB();
            string[] status = new string[3];
            if (response != null) status = parseBillAddRs(response);
            string msg = "";

            if (response != null & status[0] == "0")
            {
                msg = "Bill was added successfully!";
            }
            else
            {
                msg = "Could not add Bill " + RefNo;
                success = false;
            }

            msg = msg + "\n\n";
            msg = msg + "Status Code = " + status[0] + "\n";
            msg = msg + "Status Severity = " + status[1] + "\n";
            msg = msg + "Status Message = " + status[2] + "\n";
            //MessageBox.Show(msg);
            if (!success)
            {
                Errors = Errors + msg;
            }
            return success;
        }

        public string buildBillAddRqXML(string VendorRef, string APAccountRef, string TxDate, string RefNo, string InvLineRef, string InvLineDesc, string InvLineQuantity, string InvLineAmt, string InvLineCustomer, string InvLineClass, List<AccountingLine> lines = null )
        {

            string requestXML = "";

            //GET ALL INPUT INTO XML
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion ?? "6.1");
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement InvoiceAddRq = xmlDoc.CreateElement("BillAddRq");
            qbXMLMsgsRq.AppendChild(InvoiceAddRq);
            XmlElement InvoiceAdd = xmlDoc.CreateElement("BillAdd");
            InvoiceAddRq.AppendChild(InvoiceAdd);

            // VendorRef -> FullName
            if (!string.IsNullOrEmpty(VendorRef))
            {
                XmlElement Element_CustomerRef = xmlDoc.CreateElement("VendorRef");
                InvoiceAdd.AppendChild(Element_CustomerRef);
                XmlElement Element_CustomerRef_FullName = xmlDoc.CreateElement("ListID");
                Element_CustomerRef.AppendChild(Element_CustomerRef_FullName).InnerText = VendorRef;
            }

            // AccountRef -> FullName
            if (!string.IsNullOrEmpty(APAccountRef))
            {
                XmlElement Element_AccountRef = xmlDoc.CreateElement("APAccountRef");
                InvoiceAdd.AppendChild(Element_AccountRef);
                XmlElement Element_AccountRef_FullName = xmlDoc.CreateElement("FullName");
                Element_AccountRef.AppendChild(Element_AccountRef_FullName).InnerText = APAccountRef;
            }

            // TxnDate 
            DateTime DT_TxnDate = System.DateTime.Today;
            if (!string.IsNullOrEmpty(TxDate))
            {
                DT_TxnDate = Convert.ToDateTime(TxDate);
                string TxnDate = getDateString(DT_TxnDate);
                XmlElement Element_TxnDate = xmlDoc.CreateElement("TxnDate");
                InvoiceAdd.AppendChild(Element_TxnDate).InnerText = TxnDate;
            }

            // RefNumber 
            if (!string.IsNullOrEmpty(RefNo))
            {
                XmlElement Element_RefNumber = xmlDoc.CreateElement("RefNumber");
                InvoiceAdd.AppendChild(Element_RefNumber).InnerText = RefNo;
            }

            // Memo 
            if (!string.IsNullOrEmpty(InvLineDesc))
            {
                XmlElement Element_Memo = xmlDoc.CreateElement("Memo");
                InvoiceAdd.AppendChild(Element_Memo).InnerText = InvLineDesc;
            }

            //Item Line (REQUIRED)
            if (lines != null)
            {
                foreach (var line in lines)
                {
                    var Element_InvoiceLineAdd = xmlDoc.CreateElement("ItemLineAdd");
                    InvoiceAdd.AppendChild(Element_InvoiceLineAdd);

                    if (!string.IsNullOrEmpty(line.ItemType))
                    {
                        XmlElement Element_InvoiceLineAdd_ItemRef = xmlDoc.CreateElement("ItemRef");
                        Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_ItemRef);
                        XmlElement Element_InvoiceLineAdd_ItemRef_FullName = xmlDoc.CreateElement("FullName");
                        Element_InvoiceLineAdd_ItemRef.AppendChild(Element_InvoiceLineAdd_ItemRef_FullName).InnerText = line.ItemType;
                    }
                    if (!string.IsNullOrEmpty(line.Description))
                    {
                        XmlElement Element_InvoiceLineAdd_Desc = xmlDoc.CreateElement("Desc");
                        Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Desc).InnerText = line.Description;
                    }
                    if (!string.IsNullOrEmpty(line.Qty))
                    {
                        XmlElement Element_InvoiceLineAdd_Quantity = xmlDoc.CreateElement("Quantity");
                        Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Quantity).InnerText = line.Qty;
                    }
                    if (!string.IsNullOrEmpty(line.AmountExTax))
                    {
                        XmlElement Element_InvoiceLineAdd_Amount = xmlDoc.CreateElement("Amount");
                        Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Amount).InnerText = line.AmountExTax;
                    }
                    if (!string.IsNullOrEmpty(line.ClassRef))
                    {
                        XmlElement Element_InvoiceLineAdd_ItemRef = xmlDoc.CreateElement("ClassRef");
                        Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_ItemRef);
                        XmlElement Element_InvoiceLineAdd_ItemRef_FullName = xmlDoc.CreateElement("FullName");
                        Element_InvoiceLineAdd_ItemRef.AppendChild(Element_InvoiceLineAdd_ItemRef_FullName).InnerText = line.ClassRef;
                    }
                }
            }
            else
            {
                var Element_InvoiceLineAdd = xmlDoc.CreateElement("ItemLineAdd");
                InvoiceAdd.AppendChild(Element_InvoiceLineAdd);

                if (!string.IsNullOrEmpty(InvLineRef))
                {
                    XmlElement Element_InvoiceLineAdd_ItemRef = xmlDoc.CreateElement("ItemRef");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_ItemRef);
                    XmlElement Element_InvoiceLineAdd_ItemRef_FullName = xmlDoc.CreateElement("FullName");
                    Element_InvoiceLineAdd_ItemRef.AppendChild(Element_InvoiceLineAdd_ItemRef_FullName).InnerText = InvLineRef;
                }
                if (!string.IsNullOrEmpty(InvLineDesc))
                {
                    XmlElement Element_InvoiceLineAdd_Desc = xmlDoc.CreateElement("Desc");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Desc).InnerText = InvLineDesc;
                }
                if (!string.IsNullOrEmpty(InvLineQuantity))
                {
                    XmlElement Element_InvoiceLineAdd_Quantity = xmlDoc.CreateElement("Quantity");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Quantity).InnerText = InvLineQuantity;
                }
                if (!string.IsNullOrEmpty(InvLineAmt))
                {
                    XmlElement Element_InvoiceLineAdd_Amount = xmlDoc.CreateElement("Amount");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Amount).InnerText = InvLineAmt;
                }
                if (!string.IsNullOrEmpty(InvLineCustomer))
                {
                    XmlElement Element_InvoiceLineAdd_CustomerRef = xmlDoc.CreateElement("CustomerRef");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_CustomerRef);
                    XmlElement Element_InvoiceLineAdd_CustomerRef_FullName = xmlDoc.CreateElement("FullName");
                    Element_InvoiceLineAdd_CustomerRef.AppendChild(Element_InvoiceLineAdd_CustomerRef_FullName).InnerText = InvLineCustomer;
                }
                if (!string.IsNullOrEmpty(InvLineClass))
                {
                    XmlElement Element_InvoiceLineAdd_ClassRef = xmlDoc.CreateElement("ClassRef");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_ClassRef);
                    XmlElement Element_InvoiceLineAdd_ClassRef_FullName = xmlDoc.CreateElement("FullName");
                    Element_InvoiceLineAdd_ClassRef.AppendChild(Element_InvoiceLineAdd_ClassRef_FullName).InnerText = InvLineClass;
                }
            }


            InvoiceAddRq.SetAttribute("requestID", "99");
            requestXML = xmlDoc.OuterXml;

            return requestXML;
        }

    }
}
