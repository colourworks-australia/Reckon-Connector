using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ReckonDesktop.Repository;

namespace Reckon_Connector
{
    public partial class frmConnectWebBrowser : Form
    {
        public frmConnectWebBrowser()
        {
            InitializeComponent();
        }
        
        private void frmConnectWebBrowser_Load(object sender, EventArgs e)
        {
            //Propmt for the casdhbookId
            Uri reckonOneAuthURL = new Uri("https://identity.reckon.com/connect/authorize?client_id=" + ReckonApiHelper.DeveloperId +
            "&response_type=id_token+token&scope=openid+read+write" +
            "&redirect_uri=" + ReckonApiHelper.ReturnUrl +
            "&state=random_state&nonce=random_nonce");

            webBrowser1.Url = reckonOneAuthURL;
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            if (e.Url.ToString().StartsWith(ReckonApiHelper.ReturnUrl))
            {
                //Extract the tokens from the 

                var idt = "";
                var at = "";
                var exp = 0;
                var qs = e.Url.ToString().Replace(ReckonApiHelper.ReturnUrl + "/#", "").Split('&');

                foreach (var q in qs)
                {
                    if (q.StartsWith("id_token="))
                    {
                        idt = q.Replace("id_token=", "");
                    }
                    else if (q.StartsWith("access_token="))
                    {
                        at = q.Replace("access_token=", "");
                    }
                    else if (q.StartsWith("expires_in="))
                    {
                        exp = Convert.ToInt16(q.Replace("expires_in=", ""));
                    }
                }


                UnitOfWork _unitOfWork = new UnitOfWork();
                var existingDestination = _unitOfWork.SettingsRepository.GetByID(1);
                existingDestination.id_token = idt;
                existingDestination.access_token = at;
                var edt = DateTime.Now.AddSeconds(exp);
                existingDestination.tokenExpires = edt.Ticks;

                _unitOfWork.SettingsRepository.Detach(existingDestination);
                _unitOfWork.SettingsRepository.Update(existingDestination);
                _unitOfWork.Save();

                //Close the web browser
                this.Close();
            }
        }
    }
}
