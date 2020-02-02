using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Newtonsoft.Json;
using ReckonDesktop.Model;
using ReckonDesktop.Repository;
using ReckonDesktop.Autofile;
using System.Globalization;

namespace Reckon_Connector
{
    public partial class frmMain : Form
    {
        private static readonly UnitOfWork UnitOfWork = new UnitOfWork();
        private Settings _settings;
        private ReckonHelper RH = new ReckonHelper();
        private int CurrentTimerValue = 0;
        public string Errors = "";

        public frmMain()
        {
            InitializeComponent();
            try
            {
                _settings = UnitOfWork.SettingsRepository.GetByID(1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        
        private void cmdSettings_Click(object sender, EventArgs e)
        {
            var settingsForm = new frmSettings(FormMode.Edit, UnitOfWork, _settings.Id);
            settingsForm.ShowDialog();
            //reload
            try
            {
                var settings = UnitOfWork.SettingsRepository.Get();
                if (settings != null)
                {
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Id");
                    dt.Columns.Add("CompanyName");
                    foreach (var companyFile in settings)
                    {
                        dt.Rows.Add(companyFile.Id, companyFile.AutofileEndpoint.Replace("https://", "").Replace(".cwautofile.com.au", ""));
                    }
                    cmbSelection.DataSource = dt;
                    cmbSelection.ValueMember = "Id";
                    cmbSelection.DisplayMember = "CompanyName";
                    cmbSelection.SelectedValue = 1;
                }
                //_settings = UnitOfWork.SettingsRepository.GetByID(1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            var settings = UnitOfWork.SettingsRepository.Get();
            if (settings != null)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Id");
                dt.Columns.Add("CompanyName");
                foreach (var companyFile in settings)
                {
                    dt.Rows.Add(companyFile.Id, companyFile.AutofileEndpoint.Replace("https://","").Replace(".cwautofile.com.au",""));
                }
                cmbSelection.DataSource = dt;
                cmbSelection.ValueMember = "Id";
                cmbSelection.DisplayMember = "CompanyName";
                cmbSelection.SelectedValue = 1;
            }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        

#region CONTACTS
       
        private void button4_Click(object sender, EventArgs e)
        {
            if (_settings.UseHosted != null && _settings.UseHosted.Value == true)
            {
                GetHostedContacts();
            }
            else
            {
                GetContacts();
            }
        }
        
        private void GetContacts()
        {
            try
            {
                if (RH.connectToQB(_settings.ConnectionString ?? ""))
                {
                    string response = RH.processRequestFromQB(RH.buildCustomerQueryRqXML(new string[] { "ListID", "Name", "LastName", "FirstName", "CompanyName", "IsActive", "Notes", "ResaleNumber" }, null));
                    string response2 = RH.processRequestFromQB(RH.buildVendorQueryRqXML(new string[] { "ListID", "Name", "LastName", "FirstName", "CompanyName", "IsActive", "Notes", "VendorTaxIdent" }, null));
                    RH.disconnectFromQB();

                    var ds = new DataSet();
                    var textReader = new StringReader(response);
                    ds.ReadXml(textReader);
                    if (ds.Tables.Contains("CustomerRet"))
                    {
                        ds.Tables["CustomerRet"].Columns.Add("ContactType");
                        foreach (DataRow dr in ds.Tables["CustomerRet"].Rows)
                        {
                            dr["ContactType"] = "Customer";
                        }
                    }

                    var ds2 = new DataSet();
                    var textReader2 = new StringReader(response2);
                    ds2.ReadXml(textReader2);
                    if (ds2.Tables.Contains("VendorRet"))
                    {
                        ds2.Tables["VendorRet"].Columns.Add("ContactType");
                        foreach (DataRow dr in ds2.Tables["VendorRet"].Rows)
                        {
                            dr["ContactType"] = "Vendor";
                        }
                    }

                    if (ds.Tables.Contains("CustomerRet"))
                    {
                        ds.Tables["CustomerRet"].Merge(ds2.Tables["VendorRet"]);

                        dataGridView1.DataSource = ds.Tables["CustomerRet"];
                        var dataGridViewColumn = dataGridView1.Columns["ResaleNumber"];
                        if (dataGridViewColumn != null)
                        {
                            dataGridView1.Columns["ResaleNumber"].HeaderText = "Customer ABN";
                        }
                        var dataGridViewColumn2 = dataGridView1.Columns["VendorTaxIdent"];
                        if (dataGridViewColumn2 != null)
                        {
                            dataGridView1.Columns["VendorTaxIdent"].HeaderText = "Vendor ABN";
                        }
                    }
                    else
                    {
                        MessageBox.Show("Error: There are no customers recorded in this reckon database. Please ensure that there is at least 1 customer recorded.");
                    }
                }
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
            }
            
          
        }

        private async void GetHostedContacts()
        {
            //Authenticate and test Reckon connectivity
            if (CanBeginExecute() == false)
            {
                return;
            }

            //means we authenticated and can continue.
            try
            {
                dataGridView1.DataSource = null;
                progressBar1.Maximum = 300;
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                CurrentTimerValue = 0;
                RAHRequestTimer.Start();

                string xml1 =
                    RH.buildCustomerQueryRqXML(
                        new string[] { "ListID", "Name", "LastName", "FirstName", "CompanyName", "IsActive", "Notes" },
                        null);
                string xml2 =
                    RH.buildVendorQueryRqXML(
                        new string[] { "ListID", "Name", "LastName", "FirstName", "CompanyName", "IsActive", "Notes" },
                        null);

                var backslash = @"\";
                xml1 = xml1.Replace("\"", backslash + "\"");
                xml2 = xml2.Replace("\"", backslash + "\"");


                //TODO: Need to remove the  .Result  from the calls below to prevent the deadlock when executing the tasks so the progress bar works.
                var response = ReckonApiHelper.ExecuteRAHCall(xml1, _settings.ConnectionString, _settings.access_token).Result ?? "";
                var response2 = ReckonApiHelper.ExecuteRAHCall(xml2, _settings.ConnectionString, _settings.access_token).Result ?? "";


                if (response.StartsWith("<") && response2.StartsWith("<"))
                {
                    var ds = new DataSet();
                    var textReader = new StringReader(response);
                    ds.ReadXml(textReader);
                    ds.Tables["CustomerRet"].Columns.Add("ContactType");
                    foreach (DataRow dr in ds.Tables["CustomerRet"].Rows)
                    {
                        dr["ContactType"] = "Customer";
                    }

                    var ds2 = new DataSet();
                    var textReader2 = new StringReader(response2);
                    ds2.ReadXml(textReader2);
                    ds2.Tables["VendorRet"].Columns.Add("ContactType");
                    foreach (DataRow dr in ds2.Tables["VendorRet"].Rows)
                    {
                        dr["ContactType"] = "Vendor";
                    }


                    ds.Tables["CustomerRet"].Merge(ds2.Tables["VendorRet"]);

                    dataGridView1.DataSource = ds.Tables["CustomerRet"];
                }
                else
                {
                    //hide the rpogressbar and reset timer
                    RAHRequestTimer.Stop();
                    progressBar1.Visible = false;

                    var errMsg = "Errors Occured: " + Environment.NewLine;
                    if (!response.StartsWith("<"))
                    {
                        errMsg += response + Environment.NewLine;
                    }
                    if (!response2.StartsWith("<"))
                    {
                        errMsg += response2;
                    }
                    MessageBox.Show(errMsg);
                }
            }
            catch (Exception Ex)
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;

                if (Ex.Message != null && Ex.Message == "Could not start Reckon Accounts.")
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine +
                                    "Please ensure you have Reckon running and the company file you wish to integrate open.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }

            }
            finally
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;
            }
        }

        private async void cmdSendContactsAutofile_Click(object sender, EventArgs e)
        {
            try
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = dataGridView1.RowCount;
                progressBar1.Show();
                var cw = new AutoFile(_settings.AutofileEndpoint, _settings.AutofileUser, _settings.AutofilePassword);
                var token = await AutoFile.GetApiToken();
                DataGridViewColumnCollection Columns = dataGridView1.Columns;
                var vendorTax = false;
                var customerTax = false;
                if (Columns.Contains("VendorTaxIdent"))
                {
                    vendorTax = true;
                }
                if (Columns.Contains("ResaleNumber"))
                {
                    customerTax = true;
                }
                foreach (DataGridViewRow record in dataGridView1.Rows)
                {
                    var newContact = new AccountingContact();
                    newContact.Name = record.Cells["Name"].Value.ToString();
                    if (record.Cells["ContactType"].Value.ToString() == "Customer")
                    {
                        if (customerTax)
                        {
                            newContact.TaxNumber = record.Cells["ResaleNumber"].Value.ToString().Replace(" ", "");
                        }
                    }
                    else
                    {
                        if (vendorTax)
                        {
                            newContact.TaxNumber = record.Cells["VendorTaxIdent"].Value.ToString().Replace(" ", "");
                        }
                    }
                    newContact.TenantId = 1;
                    newContact.CardNumber = record.Cells["ListID"].Value.ToString();
                    newContact.ContactType = record.Cells["ContactType"].Value.ToString();

                    if ((newContact.Name != ""))// && (newContact.TaxNumber != ""))
                    {
                        var bills = cw.SendContacts(newContact, token);
                    }
                    progressBar1.Value = progressBar1.Value + 1;
                }
                progressBar1.Value = 0;
                progressBar1.Hide();
                MessageBox.Show("Contacts sent to autofile.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                progressBar1.Value = 0;
                progressBar1.Hide();
                MessageBox.Show("Issue: " + ex.Message, "Issue", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        
#endregion

#region ACCOUNTS

        private void button5_Click(object sender, EventArgs e)
        {
            if (_settings.UseHosted != null && _settings.UseHosted.Value == true)
            {
                GetHostedAccounts();
            }
            else
            {
                getAccounts();
            }
        }

        private void getAccounts()
        {
            try
            {
                if (RH.connectToQB(_settings.ConnectionString ?? ""))
                {
                    string response = RH.processRequestFromQB(RH.buildAccountQueryRqXML(new string[] { "ListID", "Name", "FullName", "IsActive", "AccountNumber", "AccountType" }, null));
                    //var customerList = RH.parseCustomerQueryRs(response, count);
                    RH.disconnectFromQB();

                    var ds = new DataSet();
                    var textReader = new StringReader(response);
                    ds.ReadXml(textReader);
                    dataGridView2.DataSource = ds.Tables["AccountRet"];
                }
            }
            catch (Exception Ex)
            {
                if (Ex.Message != null && Ex.Message == "Could not start Reckon Accounts.")
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine + "Please ensure you have Reckon running and the company file you wish to integrate open.");
                }
                else
                {
                    RH.disconnectFromQB();
                    MessageBox.Show("Error: " + Ex.Message);
                }

            }


        }

        private async void GetHostedAccounts()
        {
            //Authenticate and test Reckon connectivity
            if (CanBeginExecute() == false)
            {
                return;
            }

            //means we authenticated and can continue.
            try
            {
                dataGridView1.DataSource = null;
                progressBar1.Maximum = 150;
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                CurrentTimerValue = 0;
                RAHRequestTimer.Start();

                string xml1 = RH.buildAccountQueryRqXML(new string[] { "Name", "FullName", "IsActive", "AccountNumber", "AccountType" }, null);

                var backslash = @"\";
                xml1 = xml1.Replace("\"", backslash + "\"");
                //xml1 = xml1.Replace("\"", backslash + "\"");

                //TODO: Need to remove the  .Result  from the calls below to prevent the deadlock when executing the tasks so the progress bar works.
                var response = ReckonApiHelper.ExecuteRAHCall(xml1, _settings.ConnectionString, _settings.access_token).Result ?? "";
               

                if (response.StartsWith("<"))
                {
                    var ds = new DataSet();
                    var textReader = new StringReader(response);
                    ds.ReadXml(textReader);
                    dataGridView2.DataSource = ds.Tables["AccountRet"];
                }
                else
                {
                    //hide the rpogressbar and reset timer
                    RAHRequestTimer.Stop();
                    progressBar1.Visible = false;

                    var errMsg = "Errors Occured: " + Environment.NewLine + response;
                    MessageBox.Show(errMsg);
                }
            }
            catch (Exception Ex)
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;

                if (Ex.Message != null && Ex.Message == "Could not start Reckon Accounts.")
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine +
                                    "Please ensure you have Reckon running and the company file you wish to integrate open.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }

            }
            finally
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;
            }
        }


        private async void cmdSendAccountAutofile_Click(object sender, EventArgs e)
        {
            try
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = dataGridView2.RowCount;
                progressBar1.Show();
                var cw = new AutoFile(_settings.AutofileEndpoint, _settings.AutofileUser, _settings.AutofilePassword);
                var token = await AutoFile.GetApiToken();

                foreach (DataGridViewRow record in dataGridView2.Rows)
                {
                    var newContact = new AccountingAccount();
                    newContact.AccountCode = record.Cells["AccountNumber"].Value.ToString();
                    newContact.AccountDescription = record.Cells["FullName"].Value.ToString();
                    newContact.AccountId = record.Cells["ListID"].Value.ToString();
                    newContact.TenantId = 1;

                    if ((newContact.AccountCode != ""))// && (newContact.TaxNumber != ""))
                    {
                        var bills = cw.SendAccounts(newContact, token);
                    }
                    progressBar1.Value = progressBar1.Value + 1;
                }
                progressBar1.Value = 0;
                progressBar1.Hide();
                MessageBox.Show("Accounts sent to autofile.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                progressBar1.Value = 0;
                progressBar1.Hide();
                MessageBox.Show("Issue: " + ex.Message, "Issue", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }



        #endregion

#region ITEM TYPES

        private void button3_Click(object sender, EventArgs e)
        {
            //get Invoice Item Types
            if (_settings.UseHosted != null && _settings.UseHosted.Value == true)
            {
                GetHostedInvItems();
            }
            else
            {
                getInvItems();
            }
        }

        private void getInvItems()
        {
            //    var accounts = db.ExecuteQuery<Accounts>("SELECT AccountID, IsInactive, AccountName, AccountNumber FROM  Accounts").ToList();

            try
            {
                if (RH.connectToQB(_settings.ConnectionString ?? ""))
                {
                    string response = RH.processRequestFromQB(RH.buildItemQueryRqXML(new string[] { "FullName" }, null));
                    var itemList = RH.parseItemQueryRs(response);
                    RH.disconnectFromQB();

                    IList<string> listString = itemList.ToList();
                    dataGridView5.DataSource = listString.Select(x => new { FullName = x }).ToList();
                }
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
            }
        }

        private async void GetHostedInvItems()
        {
            //Authenticate and test Reckon connectivity
            if (CanBeginExecute() == false)
            {
                return;
            }

            //means we authenticated and can continue.
            try
            {
                dataGridView1.DataSource = null;
                progressBar1.Maximum = 150;
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                CurrentTimerValue = 0;
                RAHRequestTimer.Start();

                string xml1 = RH.buildItemQueryRqXML(new string[] {"FullName"}, null);

                var backslash = @"\";
                xml1 = xml1.Replace("\"", backslash + "\"");


                //TODO: Need to remove the  .Result  from the calls below to prevent the deadlock when executing the tasks so the progress bar works.
                var response = ReckonApiHelper.ExecuteRAHCall(xml1, _settings.ConnectionString, _settings.access_token).Result ?? "";


                if (response.StartsWith("<"))
                {
                    var itemList = RH.parseItemQueryRs(response);
                    IList<string> listString = itemList.ToList();
                    dataGridView5.DataSource = listString.Select(x => new { FullName = x }).ToList();
                }
                else
                {
                    //hide the rpogressbar and reset timer
                    RAHRequestTimer.Stop();
                    progressBar1.Visible = false;

                    var errMsg = "Errors Occured: " + Environment.NewLine + response;
                    MessageBox.Show(errMsg);
                }
            }
            catch (Exception Ex)
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;

                if (Ex.Message != null && Ex.Message == "Could not start Reckon Accounts.")
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine +
                                    "Please ensure you have Reckon running and the company file you wish to integrate open.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }

            }
            finally
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;
            }
        }


        private void getClasses()
        {
            //    var accounts = db.ExecuteQuery<Accounts>("SELECT AccountID, IsInactive, AccountName, AccountNumber FROM  Accounts").ToList();

            try
            {
                if (RH.connectToQB(_settings.ConnectionString ?? ""))
                {
                    string response = RH.processRequestFromQB(RH.buildClassQueryRqXML(new string[] { "ListID", "FullName" }, null));
                    //var itemList = RH.parseClassQueryRs(response);
                    RH.disconnectFromQB();

                    var ds = new DataSet();
                    var textReader = new StringReader(response);
                    ds.ReadXml(textReader);
                    dataGridView6.DataSource = ds.Tables["ClassRet"];

                    //IList<string> listString = itemList.ToList();
                    //dataGridView6.DataSource = listString.Select(x => new { FullName = x }).ToList();
                }
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
            }
        }

        private async void GetHostedClasses()
        {
            //Authenticate and test Reckon connectivity
            if (CanBeginExecute() == false)
            {
                return;
            }

            //means we authenticated and can continue.
            try
            {
                dataGridView1.DataSource = null;
                progressBar1.Maximum = 150;
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                CurrentTimerValue = 0;
                RAHRequestTimer.Start();

                string xml1 = RH.buildClassQueryRqXML(new string[] { "FullName" }, null);

                var backslash = @"\";
                xml1 = xml1.Replace("\"", backslash + "\"");


                //TODO: Need to remove the  .Result  from the calls below to prevent the deadlock when executing the tasks so the progress bar works.
                var response = ReckonApiHelper.ExecuteRAHCall(xml1, _settings.ConnectionString, _settings.access_token).Result ?? "";


                if (response.StartsWith("<"))
                {
                    var itemList = RH.parseClassQueryRs(response);
                    IList<string> listString = itemList.ToList();
                    dataGridView5.DataSource = listString.Select(x => new { FullName = x }).ToList();
                }
                else
                {
                    //hide the rpogressbar and reset timer
                    RAHRequestTimer.Stop();
                    progressBar1.Visible = false;

                    var errMsg = "Errors Occured: " + Environment.NewLine + response;
                    MessageBox.Show(errMsg);
                }
            }
            catch (Exception Ex)
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;

                if (Ex.Message != null && Ex.Message == "Could not start Reckon Accounts.")
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine +
                                    "Please ensure you have Reckon running and the company file you wish to integrate open.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }

            }
            finally
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            //Send Invoice items to Autofile
            try
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = dataGridView5.RowCount;
                progressBar1.Show();
                var cw = new AutoFile(_settings.AutofileEndpoint, _settings.AutofileUser, _settings.AutofilePassword);
                var token = await AutoFile.GetApiToken();

                foreach (DataGridViewRow record in dataGridView5.Rows)
                {
                    var newContact = new AccountingItemType();
                    newContact.TypeName = record.Cells["FullName"].Value.ToString();
                    newContact.TenantId = 1;

                    if ((newContact.TypeName != ""))// && (newContact.TaxNumber != ""))
                    {
                        var bills = cw.SendItemTypes(newContact, token);
                    }
                    progressBar1.Value = progressBar1.Value + 1;
                }
                progressBar1.Value = 0;
                progressBar1.Hide();
                MessageBox.Show("Item Types sent to autofile.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                progressBar1.Value = 0;
                progressBar1.Hide();
                MessageBox.Show("Issue: " + ex.Message, "Issue", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        #endregion




        #region Accounts Payable

        #region Add PURCHASES

        private void button2_Click(object sender, EventArgs e)
        {
            var _unitOfWork = new UnitOfWork();
            var _settings = _unitOfWork.SettingsRepository.GetByID(1);

            //used for testing
            var VendorRef = "8000003B-1308292820";
            var APAccountRef = "Accounts Payable";
            var TxDate = DateTime.Now.ToShortDateString();
            var RefNo = "ALM_REF_1";
            var InvLineRef = "Labour:RAH";
            var InvLineDesc = "Some Line Item desc";
            var InvLineQuantity = "1";
            var InvLineCustomer = "Some Line Item desc";
            var InvLineClass = "1";
            //Amount is EX GST - depending on your Reckon settings GST should be automatically added
            var InvLineAmt = "100.00";

            var res = RH.AddBillAP(_settings.ConnectionString, VendorRef, APAccountRef, TxDate, RefNo, InvLineRef, InvLineDesc, InvLineQuantity,
                InvLineAmt, InvLineCustomer, InvLineClass, ref Errors);
        }

        #endregion

        private async void button10_Click(object sender, EventArgs e)
        {
            try
            {
                var cw = new AutoFile(_settings.AutofileEndpoint, _settings.AutofileUser, _settings.AutofilePassword);

                var bills = await cw.GetBills(11);
                var so = JsonConvert.DeserializeObject<Bills>(bills);
                //Update the Grid
                dataGridView3.DataSource = so.AccountingEntries;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private async void cmdSendAPMyob_Click(object sender, EventArgs e)
        {
            //get Invoice Item Types
            if (_settings.UseHosted != null && _settings.UseHosted.Value == true)
            {
                UploadBillsHosted();
            }
            else
            {
                UploadBills();
            }
        }


        private async void UploadBills()
        {
            //NEED TO ADD InvoiceItemName to invoice table in autofile to map to a reckon item (FK to where you store them in button1_Click())
            //AND ContactID which maps to ListID from ContactsFromReckon
            //AND
            //TODO: Change so it generates a single XML file and uploads to be executed in bulk (as this is the best practice apparently)

            var _unitOfWork = new UnitOfWork();
            var _settings = _unitOfWork.SettingsRepository.GetByID(1);
            Errors = "";
            try
            {
                dataGridView3.SelectAll();
                progressBar1.Value = 0;
                progressBar1.Maximum = dataGridView3.SelectedRows.Count;
                progressBar1.Show();
                var cw = new AutoFile(_settings.AutofileEndpoint, _settings.AutofileUser, _settings.AutofilePassword);
                var token = await AutoFile.GetApiToken();

                foreach (DataGridViewRow record in dataGridView3.SelectedRows)
                {
                    //MessageBox.Show("TEST 2", "TEST");
                    var GlAccountBreakup = "";
                    var InvLineRef = "";
                    var InvLineDesc = "";
                    var InvLineCustomer = "";
                    var InvLineClass = "";
                    if (dataGridView3.Columns.Contains("GlAccountBreakup"))
                    {
                        GlAccountBreakup = (record.Cells["GlAccountBreakup"].Value ?? "").ToString();
                    }
                    //MessageBox.Show("TEST 3", "TEST");

                    var CustomerRef = record.Cells["ContactID"].Value.ToString();
                    var ARAccountRef = record.Cells["Account"].Value.ToString();
                    var TxDate = record.Cells["InvoiceDate"].Value.ToString();
                    var RefNo = record.Cells["InvoiceNumber"].Value.ToString();
                    var InvLineQuantity = "1";
                    if (dataGridView3.Columns.Contains("ItemType"))
                    {
                        InvLineRef = (record.Cells["ItemType"].Value ?? "").ToString();
                    }
                    if (dataGridView3.Columns.Contains("Description"))
                    {
                        InvLineDesc = (record.Cells["Description"].Value ?? "").ToString();
                    }
                    if (dataGridView3.Columns.Contains("JobNumber"))
                    {
                        InvLineCustomer = (record.Cells["JobNumber"].Value ?? "").ToString();
                    }
                    if (dataGridView3.Columns.Contains("Class"))
                    {
                        InvLineClass = (record.Cells["Class"].Value ?? "").ToString();
                    }
                    //Amount is EX GST - depending on your Reckon settings GST should be automatically added
                    //MessageBox.Show("TEST 4", "TEST");
                    var InvLineAmt = UtilityHelper.DblParseToDefaultValue(record.Cells["AmountExTax"].Value.ToString(), 0.00);
                    //MessageBox.Show("TEST 5", "TEST");
                    if (!string.IsNullOrEmpty(GlAccountBreakup.Replace(" ", "").Replace(",", "").Replace("|", "")))
                    {
                        List<AccountingLine> lines = new List<AccountingLine>();
                        //MessageBox.Show("TEST 6", "TEST");
                        string[] glLineSplit = GlAccountBreakup.Split('|');
                        //MessageBox.Show("TEST 7", "TEST");
                        foreach (var glLine in glLineSplit)
                        {
                            //MessageBox.Show("TEST 8", "TEST");
                            string[] glSplit = glLine.Split(',');
                            //MessageBox.Show("TEST 9", "TEST");
                            if (glSplit.Length >= 3)
                            {
                                var line = new AccountingLine();
                                //MessageBox.Show("TEST 10", "TEST");
                                var glAccount = glSplit[0];
                                //MessageBox.Show("TEST 11", "TEST");
                                var glTotal = UtilityHelper.DblParseToDefaultValue(glSplit[1], 0.00);
                                //MessageBox.Show("TEST 12", "TEST");
                                var glExTax = UtilityHelper.DblParseToDefaultValue(glSplit[2], 0.00);
                                //MessageBox.Show("TEST 13", "TEST");
                                var glDescription = "";
                                var glClass = "";
                                if (glSplit.Length >= 4)
                                {
                                    glDescription = glSplit[3];
                                    line.Description = glDescription;
                                }
                                else
                                {
                                    line.Description = InvLineDesc;
                                }
                                if (glSplit.Length >= 5)
                                {
                                    glClass = glSplit[4];
                                    line.ClassRef = glClass;
                                }

                                InvLineRef = glAccount;
                                InvLineAmt = glExTax;
                                //InvLineAmt = glTotal;
                                //MessageBox.Show("TEST 14", "TEST");
                                line.AmountExTax = glExTax.ToString("#.00", CultureInfo.InvariantCulture);
                                line.Qty = InvLineQuantity;
                                line.ItemType = InvLineRef;
                                lines.Add(line);
                                //if (RH.AddBillAP(_settings.ConnectionString, CustomerRef, ARAccountRef.ToString(), TxDate, RefNo,
                                //    InvLineRef, InvLineDesc,
                                //    InvLineQuantity,
                                //    InvLineAmt.ToString("#.00", CultureInfo.InvariantCulture), ref Errors))
                                //{
                                //    var bills = await cw.UpdateBill(UtilityHelper.IntParseToDefaultValue(record.Cells["Id"].Value.ToString(), 0), token);
                                //}
                                //MessageBox.Show("TEST 15", "TEST");
                            }
                        }
                        if (lines.Count > 0)
                        {
                            if (RH.AddBillAPLines(_settings.ConnectionString, CustomerRef, ARAccountRef.ToString(), TxDate, RefNo,
                                InvLineRef, InvLineDesc,
                                InvLineQuantity,
                                InvLineAmt.ToString("#.00", CultureInfo.InvariantCulture), InvLineCustomer, InvLineClass, lines, ref Errors))
                            {
                                var bills = await cw.UpdateBill(UtilityHelper.IntParseToDefaultValue(record.Cells["Id"].Value.ToString(), 0), token);
                            }
                        }
                    }
                    else
                    {
                        //MessageBox.Show("TEST 16", "TEST");
                        if (RH.AddBillAP(_settings.ConnectionString, CustomerRef, ARAccountRef.ToString(), TxDate, RefNo,
                            InvLineRef, InvLineDesc,
                            InvLineQuantity,
                            InvLineAmt.ToString("#.00", CultureInfo.InvariantCulture), InvLineCustomer, InvLineClass, ref Errors))
                        {
                            var bills = await cw.UpdateBill(UtilityHelper.IntParseToDefaultValue(record.Cells["Id"].Value.ToString(), 0), token);
                        }
                        //MessageBox.Show("TEST 17", "TEST");
                    }
                    //                    if (RH.AddInvoiceAR(CustomerRef, ARAccountRef.ToString(), TxDate, RefNo, InvLineRef, InvLineDesc,
                }


                progressBar1.Value = 0;
                progressBar1.Hide();
                if (Errors.Trim() != "")
                {
                    MessageBox.Show(Errors, "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                MessageBox.Show("Bills sent to Reckon.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button10_Click(null, null);
            }
            catch (Exception ex)
            {
                progressBar1.Value = 0;
                progressBar1.Hide();
                MessageBox.Show("Issue: " + ex.Message, "Issue", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private async void UploadBillsHosted()
        {
            //Authenticate and test Reckon connectivity
            if (CanBeginExecute() == false)
            {
                return;
            }

            //means we authenticated and can continue.
            try
            {
                dataGridView1.DataSource = null;
                progressBar1.Maximum = 150;
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                CurrentTimerValue = 0;
                RAHRequestTimer.Start();

                var cw = new AutoFile(_settings.AutofileEndpoint, _settings.AutofileUser, _settings.AutofilePassword);
                var token = await AutoFile.GetApiToken();

                foreach (DataGridViewRow record in dataGridView3.SelectedRows)
                {
                    var CustomerRef = record.Cells["ContactID"].Value.ToString();
                    var ARAccountRef = UtilityHelper.IntParseToDefaultValue(record.Cells["Account"].Value.ToString(), 0);
                    var TxDate = record.Cells["InvoiceDate"].Value.ToString();
                    var RefNo = record.Cells["InvoiceNumber"].Value.ToString();
                    var InvLineRef = record.Cells["InvoiceItemName"].Value.ToString() ?? "";
                    var InvLineDesc = record.Cells["Description"].Value.ToString();
                    var InvLineCustomer = record.Cells["Customer"].Value.ToString();
                    var InvLineClass = record.Cells["Class"].Value.ToString();
                    var InvLineQuantity = "1";
                    //Amount is EX GST - depending on your Reckon settings GST should be automatically added
                    var InvLineAmt = UtilityHelper.DblParseToDefaultValue(record.Cells["AmountExTax"].Value.ToString(), 0.00);

                    string xml1 = RH.buildBillAddRqXML(CustomerRef, ARAccountRef.ToString(), TxDate, RefNo, InvLineRef, InvLineDesc, InvLineQuantity, InvLineAmt.ToString("#.00", CultureInfo.InvariantCulture),InvLineCustomer, InvLineClass);

                    var backslash = @"\";
                    xml1 = xml1.Replace("\"", backslash + "\"");


                    //TODO: Need to remove the  .Result  from the calls below to prevent the deadlock when executing the tasks so the progress bar works.
                    var response = ReckonApiHelper.ExecuteRAHCall(xml1, _settings.ConnectionString, _settings.access_token).Result ?? "";


                    if (!string.IsNullOrEmpty(response) && response.StartsWith("<"))
                    {
                        string[] status = new string[3];
                        status = RH.parseBillAddRs(response);
                        var msg = "";
                        if (status != null)
                        {
                            if (status[0] == "0")
                            {
                                //Update Autofile
                                var bills = await cw.UpdateBill(UtilityHelper.IntParseToDefaultValue(record.Cells["Id"].Value.ToString(), 0), token);
                                //Add status message
                                msg = "Bill was added successfully!";
                            }
                            else
                            {
                                msg = "Could not add Bill.";
                            }

                            msg = msg + "\n\n";
                            msg = msg + "Status Code = " + status[0] + "\n";
                            msg = msg + "Status Severity = " + status[1] + "\n";
                            msg = msg + "Status Message = " + status[2] + "\n";
                        }
                        else
                        {
                            msg = "Could not add invoice.";
                        }
                    }
                    else
                    {
                        //hide the rpogressbar and reset timer
                        RAHRequestTimer.Stop();
                        progressBar1.Visible = false;

                        var errMsg = "Errors Occured: " + Environment.NewLine + response;
                        MessageBox.Show(errMsg);
                    } 
                }
            }
            catch (Exception Ex)
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;

                if (Ex.Message != null && Ex.Message == "Could not start Reckon Accounts.")
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine +
                                    "Please ensure you have Reckon running and the company file you wish to integrate open.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }

            }
            finally
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;
            }
        }



        #endregion


#region Accounts Recievable

        #region Add SALES


        private void button11_Click(object sender, EventArgs e)
        {
            var _unitOfWork = new UnitOfWork();
            var _settings = _unitOfWork.SettingsRepository.GetByID(1);


            //used for testing
            var CustomerRef = "";
            var ARAccountRef = "Accounts Receivable";
            var TxDate = DateTime.Now.ToShortDateString();
            var RefNo = "ALM_REF_1";
            var InvLineRef = "Labour:RAH";
            var InvLineDesc = "Some Line Item desc";
            var InvLineQuantity = "1";
            //Amount is EX GST - depending on your Reckon settings GST should be automatically added
            var InvLineAmt = "100.00";

            var res = RH.AddInvoiceAR(_settings.ConnectionString, CustomerRef, ARAccountRef, TxDate, RefNo, InvLineRef, InvLineDesc, InvLineQuantity,
                InvLineAmt, ref Errors);
        }


        #endregion

        private async void button9_Click(object sender, EventArgs e)
        {
            try
            {
                var cw = new AutoFile(_settings.AutofileEndpoint, _settings.AutofileUser, _settings.AutofilePassword);

                var bills = await cw.GetBills(12);
                var so = JsonConvert.DeserializeObject<Bills>(bills);
                //Update the Grid
                dataGridView4.DataSource = so.AccountingEntries;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private async void cmdSendARMyob_Click(object sender, EventArgs e)
        {
            //get Invoice Item Types
            if (_settings.UseHosted != null && _settings.UseHosted.Value == true)
            {
                UploadInvoicesHosted();
            }
            else
            {
                UploadInvoices();
            }
        }

        private async void UploadInvoices()
        {
            //NEED TO ADD InvoiceItemName to invoice table in autofile to map to a reckon item (FK to where you store them in button1_Click())
            //AND ContactID which maps to ListID from ContactsFromReckon
            var _unitOfWork = new UnitOfWork();
            var _settings = _unitOfWork.SettingsRepository.GetByID(1);
            Errors = "";

            try
            {
                dataGridView4.SelectAll();
                progressBar1.Value = 0;
                progressBar1.Maximum = dataGridView4.SelectedRows.Count;
                progressBar1.Show();
                var cw = new AutoFile(_settings.AutofileEndpoint, _settings.AutofileUser, _settings.AutofilePassword);
                var token = await AutoFile.GetApiToken();

                foreach (DataGridViewRow record in dataGridView4.SelectedRows)
                {
                    var CustomerRef = record.Cells["ContactID"].Value.ToString();
                    var ARAccountRef = record.Cells["Account"].Value.ToString();
                    var TxDate = record.Cells["InvoiceDate"].Value.ToString();
                    var RefNo = record.Cells["InvoiceNumber"].Value.ToString();
                    var InvLineRef = record.Cells["ItemType"].Value.ToString() ?? "";
                    var InvLineDesc = record.Cells["Description"].Value.ToString();
                    var InvLineQuantity = "1";
                    //Amount is EX GST - depending on your Reckon settings GST should be automatically added
                    var InvLineAmt = UtilityHelper.DblParseToDefaultValue(record.Cells["AmountExTax"].Value.ToString(), 0.00);

                    if (RH.AddInvoiceAR(_settings.ConnectionString, CustomerRef, ARAccountRef.ToString(), TxDate, RefNo, InvLineRef, InvLineDesc,
                        InvLineQuantity,
                        InvLineAmt.ToString("#.00", CultureInfo.InvariantCulture), ref Errors))
                    {
                        var bills = cw.UpdateBill(UtilityHelper.IntParseToDefaultValue(record.Cells["Id"].Value.ToString(), 0), token).Result;
                    }

                }

                progressBar1.Value = 0;
                progressBar1.Hide();
                if (Errors.Trim() != "")
                {
                    MessageBox.Show(Errors, "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                MessageBox.Show("Invoices sent to Reckon.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button9_Click(null, null);
            }
            catch (Exception ex)
            {
                progressBar1.Value = 0;
                progressBar1.Hide();
                MessageBox.Show("Issue: " + ex.Message, "Issue", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private async void UploadInvoicesHosted()
        {
            //Authenticate and test Reckon connectivity
            if (CanBeginExecute() == false)
            {
                return;
            }

            //means we authenticated and can continue.
            try
            {
                dataGridView1.DataSource = null;
                progressBar1.Maximum = 150;
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                CurrentTimerValue = 0;
                RAHRequestTimer.Start();

                var cw = new AutoFile(_settings.AutofileEndpoint, _settings.AutofileUser, _settings.AutofilePassword);
                var token = await AutoFile.GetApiToken();

                foreach (DataGridViewRow record in dataGridView4.SelectedRows)
                {
                    var CustomerRef = record.Cells["ContactID"].Value.ToString();
                    var ARAccountRef = UtilityHelper.IntParseToDefaultValue(record.Cells["Account"].Value.ToString(), 0);
                    var TxDate = record.Cells["InvoiceDate"].Value.ToString();
                    var RefNo = record.Cells["InvoiceNumber"].Value.ToString();
                    var InvLineRef = record.Cells["ItemType"].Value.ToString() ?? "";
                    var InvLineDesc = record.Cells["Description"].Value.ToString();
                    var InvLineQuantity = "1";
                    //Amount is EX GST - depending on your Reckon settings GST should be automatically added
                    var InvLineAmt = UtilityHelper.DblParseToDefaultValue(record.Cells["AmountExTax"].Value.ToString(), 0.00);
                    
                    string xml1 = RH.buildInvoiceAddRqXML(CustomerRef, ARAccountRef.ToString(), TxDate, RefNo, InvLineRef, InvLineDesc, InvLineQuantity, InvLineAmt.ToString("#.00", CultureInfo.InvariantCulture));

                    var backslash = @"\";
                    xml1 = xml1.Replace("\"", backslash + "\"");


                    //TODO: Need to remove the  .Result  from the calls below to prevent the deadlock when executing the tasks so the progress bar works.
                    var response = ReckonApiHelper.ExecuteRAHCall(xml1, _settings.ConnectionString, _settings.access_token).Result ?? "";


                    if (!string.IsNullOrEmpty(response) && response.StartsWith("<"))
                    {
                        string[] status = new string[3];
                        status = RH.parseBillAddRs(response);
                        var msg = "";
                        if (status != null)
                        {
                            if (status[0] == "0")
                            {
                                //Update AUtofile
                                var bills = cw.UpdateBill(UtilityHelper.IntParseToDefaultValue(record.Cells["Id"].Value.ToString(), 0), token).Result;
                                //Add status messgae
                                msg = "Invoice was added successfully!";
                            }
                            else
                            {
                                msg = "Could not add invoice.";
                            }

                            msg = msg + "\n\n";
                            msg = msg + "Status Code = " + status[0] + "\n";
                            msg = msg + "Status Severity = " + status[1] + "\n";
                            msg = msg + "Status Message = " + status[2] + "\n";
                        }
                        else
                        {
                            msg = "Could not add invoice.";
                        }
                    }
                    else
                    {
                        //hide the rpogressbar and reset timer
                        RAHRequestTimer.Stop();
                        progressBar1.Visible = false;

                        var errMsg = "Errors Occured: " + Environment.NewLine + response;
                        MessageBox.Show(errMsg);
                    }
                }
            }
            catch (Exception Ex)
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;

                if (Ex.Message != null && Ex.Message == "Could not start Reckon Accounts.")
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine +
                                    "Please ensure you have Reckon running and the company file you wish to integrate open.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }

            }
            finally
            {
                //hide the rpogressbar and reset timer
                RAHRequestTimer.Stop();
                progressBar1.Visible = false;
            }
        }

        #endregion

        private void button6_Click(object sender, EventArgs e)
        {
            var nfrm = new frmConnectWebBrowser();
            nfrm.ShowDialog();
        }

        private void RAHRequestTimer_Tick(object sender, EventArgs e)
        {
            CurrentTimerValue += 1;
            if (CurrentTimerValue < progressBar1.Maximum)
            {
                progressBar1.Value = CurrentTimerValue;
            }
            
        }

        public static bool CanBeginExecute()
        {
            try
            {

                var _unitOfWork = new UnitOfWork();
                var _settings = _unitOfWork.SettingsRepository.GetByID(1);

                
                if (_settings.tokenExpires == null || _settings.tokenExpires.Value <= DateTime.Now.Ticks)
                {
                    //means auth is required
                    var nfrm = new frmConnectWebBrowser();
                    nfrm.ShowDialog();

                    //reload the settings
                    _unitOfWork = new UnitOfWork();
                    _settings = _unitOfWork.SettingsRepository.GetByID(1);

                    //If expiry still out then ABORT
                    if (_settings.tokenExpires != null && _settings.tokenExpires.Value <= DateTime.Now.Ticks)
                    {
                        MessageBox.Show("Unable to refresh your authentication with Reckon please try again.");
                        return false;
                    }

                    //warm up the connection to the Hosted API
                    var res = ReckonApiHelper.ExecuteCallHeartBeat(_settings.access_token).Result;
                    if (res != "OK")
                    {
                        MessageBox.Show("Reckon API service is unreachable. Check your internet connection and whether the Reckon Service is running.");
                        return false;
                    }
                }

            }
            catch (Exception ex1)
            {
                MessageBox.Show("Error: " + ex1.Message);
                return false;
            }

            //Otherwise return true
            return true;
        }

        private void cmdGetClasses_Click(object sender, EventArgs e)
        {
            //get Classes
            if (_settings.UseHosted != null && _settings.UseHosted.Value == true)
            {
                GetHostedClasses();
            }
            else
            {
                getClasses();
            }
        }

        private async void cmdSendClasses_Click(object sender, EventArgs e)
        {
            try
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = dataGridView6.RowCount;
                progressBar1.Show();
                var cw = new AutoFile(_settings.AutofileEndpoint, _settings.AutofileUser, _settings.AutofilePassword);
                var token = await AutoFile.GetApiToken();

                foreach (DataGridViewRow record in dataGridView6.Rows)
                {
                    var newContact = new AccountingJob();
                    newContact.JobId = record.Cells["ListID"].Value.ToString();
                    newContact.JobNumber = record.Cells["FullName"].Value.ToString();
                    newContact.JobDescription = record.Cells["FullName"].Value.ToString();
                    newContact.JobName = record.Cells["FullName"].Value.ToString();
                    newContact.ContactName = record.Cells["FullName"].Value.ToString();
                    newContact.TenantId = 1;

                    if ((newContact.JobNumber != ""))// && (newContact.TaxNumber != ""))
                    {
                        var bills = cw.SendJobs(newContact, token);
                    }
                    progressBar1.Value = progressBar1.Value + 1;
                }
                progressBar1.Value = 0;
                progressBar1.Hide();
                MessageBox.Show("Classes sent to autofile.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                progressBar1.Value = 0;
                progressBar1.Hide();
                MessageBox.Show("Issue: " + ex.Message, "Issue", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void cmbSelection_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ComboBox cmb = (ComboBox)sender;
                DataRowView drv = (DataRowView)cmb.SelectedItem;
                var selectedValue = drv.Row.ItemArray[0];
                var intSelectedValue = UtilityHelper.IntParseToNull(selectedValue.ToString());
                if (intSelectedValue != null)
                {
                    var setting = UnitOfWork.SettingsRepository.GetByID(intSelectedValue);
                    if (setting != null)
                    {
                        _settings = setting;
                    }
                }

            }
            catch (Exception exception)
            {
            }
        }
    }
}