using ReckonDesktop.Model;
using System;
using System.Data.Odbc;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Reckon_Connector
{
    public static class DbInteraction
    {
        public static string LogFileName = "C:\\Premier19\\MYOBPLOG.TXT";

        public static bool AddMiscPurchase(ImportMiscellaneousPurchase rec, string connString)
        {
            try
            {

                var cols = "";
                var vals = "";

                if (!string.IsNullOrEmpty(rec.CoLastName))
                {
                    cols += ",CoLastName";
                    vals += ",'" + rec.CoLastName + "'";
                }

                if (!string.IsNullOrEmpty(rec.FirstName))
                {
                    cols += ",FirstName";
                    vals += ",'" + rec.FirstName + "'";
                }

                if (!string.IsNullOrEmpty(rec.PurchaseNumber))
                {
                    cols += ",PurchaseNumber";
                    vals += ",'" + rec.PurchaseNumber + "'";
                }
                if (!string.IsNullOrEmpty(rec.PurchaseDate))
                {
                    cols += ",PurchaseDate";
                    vals += ",'" + rec.PurchaseDate + "'";
                }

                if (!string.IsNullOrEmpty(rec.SuppliersNumber))
                {
                    cols += ",SuppliersNumber";
                    vals += ",'" + rec.SuppliersNumber + "'";
                }

                if (!string.IsNullOrEmpty(rec.Inclusive))
                {
                    cols += ",Inclusive";
                    vals += ",'" + rec.Inclusive + "'";
                }

                if (!string.IsNullOrEmpty(rec.Memo))
                {
                    cols += ",Memo";
                    vals += ",'" + rec.Memo + "'";
                }

                if (!string.IsNullOrEmpty(rec.Description))
                {
                    cols += ",Description";
                    vals += ",'" + rec.Description + "'";
                }

                if (!string.IsNullOrEmpty(rec.Job))
                {
                    cols += ",Job";
                    vals += ",'" + rec.Job + "'";
                }

                if (!string.IsNullOrEmpty(rec.TaxCode))
                {
                    cols += ",TaxCode";
                    vals += ",'" + rec.TaxCode + "'";
                }

                if (!string.IsNullOrEmpty(rec.CurrencyCode))
                {
                    cols += ",CurrencyCode";
                    vals += ",'" + rec.CurrencyCode + "'";
                }

                if (!string.IsNullOrEmpty(rec.Category))
                {
                    cols += ",Category";
                    vals += ",'" + rec.Category + "'";
                }

                if (!string.IsNullOrEmpty(rec.CardID))
                {
                    cols += ",CardID";
                    vals += ",'" + rec.CardID + "'";
                }


                //DOubles and integers

                if (rec.NonGSTImportAmount != null)
                {
                    cols += ",NonGSTImportAmount";
                    vals += "," + rec.NonGSTImportAmount.ToString();
                }

                if (rec.GSTAmount != null)
                {
                    cols += ",GSTAmount";
                    vals += "," + rec.GSTAmount.ToString();
                }

                if (rec.ImportDutyAmount != null)
                {
                    cols += ",ImportDutyAmount";
                    vals += "," + rec.ImportDutyAmount.ToString();
                }

                if (rec.ExTaxAmount != null)
                {
                    cols += ",ExTaxAmount";
                    vals += "," + rec.ExTaxAmount.ToString();
                }

                if (rec.IncTaxAmount != null)
                {
                    cols += ",IncTaxAmount";
                    vals += "," + rec.IncTaxAmount.ToString();
                }

                if (rec.ExchangeRate != null)
                {
                    cols += ",ExchangeRate";
                    vals += "," + rec.ExchangeRate.ToString();
                }

                if (rec.PercentDiscount != null)
                {
                    cols += ",PercentDiscount";
                    vals += "," + rec.PercentDiscount.ToString();
                }

                if (rec.AmountPaid != null)
                {
                    cols += ",AmountPaid";
                    vals += "," + rec.AmountPaid.ToString();
                }

                if (rec.RecordID != null)
                {
                    cols += ",RecordID";
                    vals += "," + rec.RecordID.ToString();
                }

                if (rec.AccountNumber != null)
                {
                    cols += ",AccountNumber";
                    vals += "," + rec.AccountNumber.ToString();
                }

                if (rec.PaymentIsDue != null)
                {
                    cols += ",PaymentIsDue";
                    vals += "," + rec.PaymentIsDue.ToString();
                }

                if (rec.DiscountDays != null)
                {
                    cols += ",DiscountDays";
                    vals += "," + rec.DiscountDays.ToString();
                }

                if (rec.BalanceDueDays != null)
                {
                    cols += ",BalanceDueDays";
                    vals += "," + rec.BalanceDueDays.ToString();
                }

                //remove first comma
                cols = cols.Substring(1, cols.Length - 1);
                vals = vals.Substring(1, vals.Length - 1);

                var sqlquery = "Insert into Import_Miscellaneous_Purchases (" + cols + ") values (" + vals + ")";


                //Delete the LOG File - Probably not required, but ensures its clean
                File.Delete(LogFileName);

                using (var connection = new OdbcConnection(connString))
                {
                    var db = new dbLinqDataContext(connection);

                    var res = db.ExecuteCommand(sqlquery);

                    if (res < 1)
                    {
                        return false;
                    }
                }

                //Read the log file
                var sr = File.ReadAllText(LogFileName);

                if (sr.StartsWith("-"))
                {
                    var errMsg = "An Unknown Validation Error has occured while importing, check the log file (" +
                                   LogFileName + ") for details.";
                    //means there was an issue
                    var newLinesRegex = new Regex(@"\r\n|\n|\r", RegexOptions.Singleline);
                    var lines = newLinesRegex.Split(sr);
                    foreach (string l in lines)
                    {
                        if (l.StartsWith("Error -"))
                        {
                            errMsg = l;
                        }
                    }

                    MessageBox.Show("MYOB has rejected the import with the following reason: " + errMsg);
                    return false;
                }


            }
            catch (Exception Ex)
            {
                //If you're getting an error about the file already being opened exclusivley - you HAVE to run as Administrator.
                if (Ex.Message != null && Ex.Message.Contains("[HZ080]"))
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine + "Please ensure you have MYOB running, that you are running this program and MYOB as an administrator (Windows 7+) That you have not exceeded your licenced User count.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }
                return false;
            }

            return true;
        }

        public static bool AddMiscSale(ImportMiscellaneousSale rec, string connString)
        {
            try
            {

                var cols = "";
                var vals = "";

                if (!string.IsNullOrEmpty(rec.CoLastName))
                {
                    cols += ",CoLastName";
                    vals += ",'" + rec.CoLastName + "'";
                }

                if (!string.IsNullOrEmpty(rec.FirstName))
                {
                    cols += ",FirstName";
                    vals += ",'" + rec.FirstName + "'";
                }

                if (!string.IsNullOrEmpty(rec.CustomersNumber))
                {
                    cols += ",CustomersNumber";
                    vals += ",'" + rec.CustomersNumber + "'";
                }
                if (!string.IsNullOrEmpty(rec.SaleDate))
                {
                    cols += ",SaleDate";
                    vals += ",'" + rec.SaleDate + "'";
                }

                if (!string.IsNullOrEmpty(rec.InvoiceNumber))
                {
                    cols += ",InvoiceNumber";
                    vals += ",'" + rec.InvoiceNumber + "'";
                }

                if (!string.IsNullOrEmpty(rec.Inclusive))
                {
                    cols += ",Inclusive";
                    vals += ",'" + rec.Inclusive + "'";
                }

                if (!string.IsNullOrEmpty(rec.Memo))
                {
                    cols += ",Memo";
                    vals += ",'" + rec.Memo + "'";
                }

                if (!string.IsNullOrEmpty(rec.Description))
                {
                    cols += ",Description";
                    vals += ",'" + rec.Description + "'";
                }

                if (!string.IsNullOrEmpty(rec.Job))
                {
                    cols += ",Job";
                    vals += ",'" + rec.Job + "'";
                }

                if (!string.IsNullOrEmpty(rec.TaxCode))
                {
                    cols += ",TaxCode";
                    vals += ",'" + rec.TaxCode + "'";
                }

                if (!string.IsNullOrEmpty(rec.CurrencyCode))
                {
                    cols += ",CurrencyCode";
                    vals += ",'" + rec.CurrencyCode + "'";
                }

                if (!string.IsNullOrEmpty(rec.Category))
                {
                    cols += ",Category";
                    vals += ",'" + rec.Category + "'";
                }

                if (!string.IsNullOrEmpty(rec.CardID))
                {
                    cols += ",CardID";
                    vals += ",'" + rec.CardID + "'";
                }

                if (!string.IsNullOrEmpty(rec.SalespersonLastName))
                {
                    cols += ",SalespersonLastName";
                    vals += ",'" + rec.SalespersonLastName + "'";
                }

                if (!string.IsNullOrEmpty(rec.SalespersonFirstName))
                {
                    cols += ",SalespersonFirstName";
                    vals += ",'" + rec.SalespersonFirstName + "'";
                }

                if (!string.IsNullOrEmpty(rec.ReferralSource))
                {
                    cols += ",ReferralSource";
                    vals += ",'" + rec.ReferralSource + "'";
                }

                if (!string.IsNullOrEmpty(rec.SaleStatus))
                {
                    cols += ",SaleStatus";
                    vals += ",'" + rec.SaleStatus + "'";
                }

                if (!string.IsNullOrEmpty(rec.PaymentMethod))
                {
                    cols += ",PaymentMethod";
                    vals += ",'" + rec.PaymentMethod + "'";
                }

                if (!string.IsNullOrEmpty(rec.PaymentNotes))
                {
                    cols += ",PaymentNotes";
                    vals += ",'" + rec.PaymentNotes + "'";
                }

                if (!string.IsNullOrEmpty(rec.NameOnCard))
                {
                    cols += ",NameOnCard";
                    vals += ",'" + rec.NameOnCard + "'";
                }

                if (!string.IsNullOrEmpty(rec.CardNumber))
                {
                    cols += ",CardNumber";
                    vals += ",'" + rec.CardNumber + "'";
                }

                if (!string.IsNullOrEmpty(rec.ExpiryDate))
                {
                    cols += ",ExpiryDate";
                    vals += ",'" + rec.ExpiryDate + "'";
                }

                if (!string.IsNullOrEmpty(rec.AuthorisationCode))
                {
                    cols += ",AuthorisationCode";
                    vals += ",'" + rec.AuthorisationCode + "'";
                }

                if (!string.IsNullOrEmpty(rec.DrawerBSB))
                {
                    cols += ",DrawerBSB";
                    vals += ",'" + rec.DrawerBSB + "'";
                }

                if (!string.IsNullOrEmpty(rec.DrawerAccountName))
                {
                    cols += ",DrawerAccountName";
                    vals += ",'" + rec.DrawerAccountName + "'";
                }

                if (!string.IsNullOrEmpty(rec.DrawerChequeNumber))
                {
                    cols += ",DrawerChequeNumber";
                    vals += ",'" + rec.DrawerChequeNumber + "'";
                }

                if (!string.IsNullOrEmpty(rec.Category))
                {
                    cols += ",Category";
                    vals += ",'" + rec.Category + "'";
                }





                //DOubles and integers

                if (rec.NonGSTLCTAmount != null)
                {
                    cols += ",NonGSTLCTAmount";
                    vals += "," + rec.NonGSTLCTAmount.ToString();
                }

                if (rec.GSTAmount != null)
                {
                    cols += ",GSTAmount";
                    vals += "," + rec.GSTAmount.ToString();
                }

                if (rec.LCTAmount != null)
                {
                    cols += ",LCTAmount";
                    vals += "," + rec.LCTAmount.ToString();
                }

                if (rec.ExTaxAmount != null)
                {
                    cols += ",ExTaxAmount";
                    vals += "," + rec.ExTaxAmount.ToString();
                }

                if (rec.IncTaxAmount != null)
                {
                    cols += ",IncTaxAmount";
                    vals += "," + rec.IncTaxAmount.ToString();
                }

                if (rec.ExchangeRate != null)
                {
                    cols += ",ExchangeRate";
                    vals += "," + rec.ExchangeRate.ToString();
                }

                if (rec.PercentDiscount != null)
                {
                    cols += ",PercentDiscount";
                    vals += "," + rec.PercentDiscount.ToString();
                }

                if (rec.PercentMonthlyCharge != null)
                {
                    cols += ",PercentMonthlyCharge";
                    vals += "," + rec.PercentMonthlyCharge.ToString();
                }

                if (rec.AmountPaid != null)
                {
                    cols += ",AmountPaid";
                    vals += "," + rec.AmountPaid.ToString();
                }

                if (rec.RecordID != null)
                {
                    cols += ",RecordID";
                    vals += "," + rec.RecordID.ToString();
                }

                if (rec.AccountNumber != null)
                {
                    cols += ",AccountNumber";
                    vals += "," + rec.AccountNumber.ToString();
                }

                if (rec.PaymentIsDue != null)
                {
                    cols += ",PaymentIsDue";
                    vals += "," + rec.PaymentIsDue.ToString();
                }

                if (rec.DiscountDays != null)
                {
                    cols += ",DiscountDays";
                    vals += "," + rec.DiscountDays.ToString();
                }

                if (rec.BalanceDueDays != null)
                {
                    cols += ",BalanceDueDays";
                    vals += "," + rec.BalanceDueDays.ToString();
                }

                if (rec.DrawerAccountNumber != null)
                {
                    cols += ",DrawerAccountNumber";
                    vals += "," + rec.DrawerAccountNumber.ToString();
                }


                //remove first comma
                cols = cols.Substring(1, cols.Length - 1);
                vals = vals.Substring(1, vals.Length - 1);

                var sqlquery = "Insert into Import_Miscellaneous_Sales (" + cols + ") values (" + vals + ")";


                //Delete the LOG File - Probably not required, but ensures its clean
                File.Delete(LogFileName);

                using (var connection = new OdbcConnection(connString))
                {
                    var db = new dbLinqDataContext(connection);

                    var res = db.ExecuteCommand(sqlquery);

                    if (res < 1)
                    {
                        return false;
                    }
                }

                //Read the log file
                var sr = File.ReadAllText(LogFileName);

                if (sr.StartsWith("-"))
                {
                    var errMsg = "An Unknown Validation Error has occured while importing, check the log file (" +
                                   LogFileName + ") for details.";
                    //means there was an issue
                    var newLinesRegex = new Regex(@"\r\n|\n|\r", RegexOptions.Singleline);
                    var lines = newLinesRegex.Split(sr);
                    foreach (string l in lines)
                    {
                        if (l.StartsWith("Error -"))
                        {
                            errMsg = l;
                        }
                    }

                    MessageBox.Show("MYOB has rejected the import with the following reason: " + errMsg);
                    return false;
                }


            }
            catch (Exception Ex)
            {
                //If you're getting an error about the file already being opened exclusivley - you HAVE to run as Administrator.
                if (Ex.Message != null && Ex.Message.Contains("[HZ080]"))
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine + "Please ensure you have MYOB running, that you are running this program and MYOB as an administrator (Windows 7+) That you have not exceeded your licenced User count.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }
                return false;
            }

            return true;
        }

        public static bool AddSupplier(ImportSupplierCard rec, string connString)
        {
            try
            {

                var cols = "";
                var vals = "";

                if (!string.IsNullOrEmpty(rec.CoLastName))
                {
                    cols += ",CoLastName";
                    vals += ",'" + rec.CoLastName + "'";
                }

                if (!string.IsNullOrEmpty(rec.FirstName))
                {
                    cols += ",FirstName";
                    vals += ",'" + rec.FirstName + "'";
                }

                if (!string.IsNullOrEmpty(rec.CardID))
                {
                    cols += ",CardID";
                    vals += ",'" + rec.CardID + "'";
                }

                if (!string.IsNullOrEmpty(rec.CardStatus))
                {
                    cols += ",CardStatus";
                    vals += ",'" + rec.CardStatus + "'";
                }

                if (!string.IsNullOrEmpty(rec.CurrencyCode))
                {
                    cols += ",CurrencyCode";
                    vals += ",'" + rec.CurrencyCode + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1AddressLine1))
                {
                    cols += ",Address1AddressLine1";
                    vals += ",'" + rec.Address1AddressLine1 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1AddressLine2))
                {
                    cols += ",Address1AddressLine2";
                    vals += ",'" + rec.Address1AddressLine2 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1AddressLine3))
                {
                    cols += ",Address1AddressLine3";
                    vals += ",'" + rec.Address1AddressLine3 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1AddressLine4))
                {
                    cols += ",Address1AddressLine4";
                    vals += ",'" + rec.Address1AddressLine4 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1City))
                {
                    cols += ",Address1City";
                    vals += ",'" + rec.Address1City + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1State))
                {
                    cols += ",Address1State";
                    vals += ",'" + rec.Address1State + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1PostCode))
                {
                    cols += ",Address1PostCode";
                    vals += ",'" + rec.Address1PostCode + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Country))
                {
                    cols += ",Address1Country";
                    vals += ",'" + rec.Address1Country + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Phone1))
                {
                    cols += ",Address1Phone1";
                    vals += ",'" + rec.Address1Phone1 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Phone2))
                {
                    cols += ",Address1Phone2";
                    vals += ",'" + rec.Address1Phone2 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Phone3))
                {
                    cols += ",Address1Phone3";
                    vals += ",'" + rec.Address1Phone3 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Fax))
                {
                    cols += ",Address1Fax";
                    vals += ",'" + rec.Address1Fax + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Email))
                {
                    cols += ",Address1Email";
                    vals += ",'" + rec.Address1Email + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Website))
                {
                    cols += ",Address1Website";
                    vals += ",'" + rec.Address1Website + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1ContactName))
                {
                    cols += ",Address1ContactName";
                    vals += ",'" + rec.Address1ContactName + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Salutation))
                {
                    cols += ",Address1Salutation";
                    vals += ",'" + rec.Address1Salutation + "'";
                }





                //remove first comma
                cols = cols.Substring(1, cols.Length - 1);
                vals = vals.Substring(1, vals.Length - 1);

                var sqlquery = "Insert into Import_Supplier_Cards (" + cols + ") values (" + vals + ")";


                //Delete the LOG File - Probably not required, but ensures its clean
                File.Delete(LogFileName);

                using (var connection = new OdbcConnection(connString))
                {
                    var db = new dbLinqDataContext(connection);

                    var res = db.ExecuteCommand(sqlquery);

                    if (res < 1)
                    {
                        return false;
                    }
                }

                //Read the log file
                var sr = File.ReadAllText(LogFileName);

                if (sr.StartsWith("-"))
                {
                    var errMsg = "An Unknown Validation Error has occured while importing, check the log file (" +
                                   LogFileName + ") for details.";
                    //means there was an issue
                    var newLinesRegex = new Regex(@"\r\n|\n|\r", RegexOptions.Singleline);
                    var lines = newLinesRegex.Split(sr);
                    foreach (string l in lines)
                    {
                        if (l.StartsWith("Error -"))
                        {
                            errMsg = l;
                        }
                    }

                    MessageBox.Show("MYOB has rejected the import with the following reason: " + errMsg);
                    return false;
                }


            }
            catch (Exception Ex)
            {
                //If you're getting an error about the file already being opened exclusivley - you HAVE to run as Administrator.
                if (Ex.Message != null && Ex.Message.Contains("[HZ080]"))
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine + "Please ensure you have MYOB running, that you are running this program and MYOB as an administrator (Windows 7+) That you have not exceeded your licenced User count.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }
                return false;
            }

            return true;
        }

        public static bool AddCustomer(ImportCustomerCard rec, string connString)
        {
            try
            {

                var cols = "";
                var vals = "";


                if (!string.IsNullOrEmpty(rec.CoLastName))
                {
                    cols += ",CoLastName";
                    vals += ",'" + rec.CoLastName + "'";
                }

                if (!string.IsNullOrEmpty(rec.FirstName))
                {
                    cols += ",FirstName";
                    vals += ",'" + rec.FirstName + "'";
                }

                if (!string.IsNullOrEmpty(rec.CardID))
                {
                    cols += ",CardID";
                    vals += ",'" + rec.CardID + "'";
                }

                if (!string.IsNullOrEmpty(rec.CardStatus))
                {
                    cols += ",CardStatus";
                    vals += ",'" + rec.CardStatus + "'";
                }

                if (!string.IsNullOrEmpty(rec.CurrencyCode))
                {
                    cols += ",CurrencyCode";
                    vals += ",'" + rec.CurrencyCode + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1AddressLine1))
                {
                    cols += ",Address1AddressLine1";
                    vals += ",'" + rec.Address1AddressLine1 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1AddressLine2))
                {
                    cols += ",Address1AddressLine2";
                    vals += ",'" + rec.Address1AddressLine2 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1AddressLine3))
                {
                    cols += ",Address1AddressLine3";
                    vals += ",'" + rec.Address1AddressLine3 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1AddressLine4))
                {
                    cols += ",Address1AddressLine4";
                    vals += ",'" + rec.Address1AddressLine4 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1City))
                {
                    cols += ",Address1City";
                    vals += ",'" + rec.Address1City + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1State))
                {
                    cols += ",Address1State";
                    vals += ",'" + rec.Address1State + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1PostCode))
                {
                    cols += ",Address1PostCode";
                    vals += ",'" + rec.Address1PostCode + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Country))
                {
                    cols += ",Address1Country";
                    vals += ",'" + rec.Address1Country + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Phone1))
                {
                    cols += ",Address1Phone1";
                    vals += ",'" + rec.Address1Phone1 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Phone2))
                {
                    cols += ",Address1Phone2";
                    vals += ",'" + rec.Address1Phone2 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Phone3))
                {
                    cols += ",Address1Phone3";
                    vals += ",'" + rec.Address1Phone3 + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Fax))
                {
                    cols += ",Address1Fax";
                    vals += ",'" + rec.Address1Fax + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Email))
                {
                    cols += ",Address1Email";
                    vals += ",'" + rec.Address1Email + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Website))
                {
                    cols += ",Address1Website";
                    vals += ",'" + rec.Address1Website + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1ContactName))
                {
                    cols += ",Address1ContactName";
                    vals += ",'" + rec.Address1ContactName + "'";
                }

                if (!string.IsNullOrEmpty(rec.Address1Salutation))
                {
                    cols += ",Address1Salutation";
                    vals += ",'" + rec.Address1Salutation + "'";
                }







                //remove first comma
                cols = cols.Substring(1, cols.Length - 1);
                vals = vals.Substring(1, vals.Length - 1);

                var sqlquery = "Insert into Import_Customer_Cards (" + cols + ") values (" + vals + ")";


                //Delete the LOG File - Probably not required, but ensures its clean
                File.Delete(LogFileName);

                using (var connection = new OdbcConnection(connString))
                {
                    var db = new dbLinqDataContext(connection);

                    var res = db.ExecuteCommand(sqlquery);

                    if (res < 1)
                    {
                        return false;
                    }
                }

                //Read the log file
                var sr = File.ReadAllText(LogFileName);

                if (sr.StartsWith("-"))
                {
                    var errMsg = "An Unknown Validation Error has occured while importing, check the log file (" +
                                   LogFileName + ") for details.";
                    //means there was an issue
                    var newLinesRegex = new Regex(@"\r\n|\n|\r", RegexOptions.Singleline);
                    var lines = newLinesRegex.Split(sr);
                    foreach (string l in lines)
                    {
                        if (l.StartsWith("Error -"))
                        {
                            errMsg = l;
                        }
                    }

                    MessageBox.Show("MYOB has rejected the import with the following reason: " + errMsg);
                    return false;
                }


            }
            catch (Exception Ex)
            {
                //If you're getting an error about the file already being opened exclusivley - you HAVE to run as Administrator.
                if (Ex.Message != null && Ex.Message.Contains("[HZ080]"))
                {
                    MessageBox.Show("Error: " + Ex.Message + Environment.NewLine + Environment.NewLine + "Please ensure you have MYOB running, that you are running this program and MYOB as an administrator (Windows 7+) That you have not exceeded your licenced User count.");
                }
                else
                {
                    MessageBox.Show("Error: " + Ex.Message);
                }
                return false;
            }

            return true;
        }

    }
}
