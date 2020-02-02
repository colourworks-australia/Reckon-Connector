using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Newtonsoft.Json;

using System.Web;

namespace Reckon_Connector
{
    public static class ReckonApiHelper
    {
        //This one should stay the same
        public const string IdentityServerUrl = "https://identity.reckon.com/connect/authorize";

        //You will need to change this to the AutoFile guys developer account Security ID and Sub Key
        public const string DeveloperId = "79f4ac48-3fc6-435f-a6b9-2dd7a1487505"; //This is used to call the identity server
        public const string SubscriptionKey = "ec97676da3c44bb492fc77d43828b030"; //This is used in ALL calls to the API

        //The ReturnURL needs to be changed to match whatever you get the Reckon Guys to change it to....
        public const string ReturnUrl = "http://localhost";
        //public const string ReturnUrl = "https://internal.cwautofile.com.au";
        public const string ClientServerVersion = "2015.R2.AU";

        public static void PromptForoAuth()
        {
            var nfrm = new frmConnectWebBrowser();
            nfrm.ShowDialog();
        }

      
        public static async Task<string> ExecuteCallHeartBeat(string access_token)
        {
            //get the list of companies we have access to:
            var client = new HttpClient();

            // Request headers
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", ReckonApiHelper.SubscriptionKey);
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + access_token);

            var uri = "https://api.reckon.com/RAH/Heartbeat";
            var response = client.GetAsync(uri).Result;

            return response.StatusCode.ToString();
            
        }

        public static async Task<string> ExecuteRAHCall(string formedQBXMLRequest, string filename, string access_token,
            string querystring = "")
        {
            try
            {

                var client = new HttpClient();
                //this is important gives the reckon server time to return SOMETHING....
                TimeSpan HttpTimeSpan = new TimeSpan(0, 0, 3, 0);
                client.Timeout = (HttpTimeSpan);

                var customJSON = "{" + Environment.NewLine + "\"FileName\"" + ":" + "\"" + filename + "\"" + "," + Environment.NewLine +
                                 "\"Operation\"" + ":" + "\"" + formedQBXMLRequest + "\"" + "," + Environment.NewLine + 
                                 "\"CountryVersion\"" + ":" + "\"" + ClientServerVersion + "\"" + Environment.NewLine + "}";

                // Request headers
                client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", ReckonApiHelper.SubscriptionKey);
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + access_token);


                var uri = "https://api.reckon.com/RAH/v2/?" + querystring;

                HttpResponseMessage response;

                // Request body
                byte[] byteData = Encoding.UTF8.GetBytes(customJSON);

                using (var content = new ByteArrayContent(byteData))
                {
                    content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                    response = client.PostAsync(uri, content).Result;
                }

                //I don't know why but fuck me if this doesn't work to fix the Cannot Connect error....
                System.Threading.Thread.Sleep(2000);

                //Check the response status code and make sure that its sucessful
                var result = response.Content.ReadAsStringAsync().Result;

                if (!string.IsNullOrEmpty(result))
                {
                    if (result.StartsWith("[") || result.StartsWith("{"))
                    {
                        Dictionary<string, string> values = JsonConvert.DeserializeObject<Dictionary<string, string>>(result);
                        if (values.ContainsKey("RequestId"))
                        {
                            if (values["RetryLater"] == "True")
                            {
                                //means we need to retry again (later)
                                DialogResult rez = MessageBox.Show("Your Data is still processing on the Reckon Server. You can continue to wait or try again later. Do you want to continue to wait?", "Confirmation", MessageBoxButtons.YesNoCancel);
                                if (rez == DialogResult.Yes)
                                {
                                    //recursive call to the retrieve function to continue waiting
                                    return RetrieveExecutedRAHCall(values["RequestId"], access_token).Result;
                                }
                            }
                            else
                            {
                                //return the XML payload
                                return values["Data"];
                            }
                        }
                        else if (values.ContainsKey("Message"))
                        {
                            //probably some error or something so return the message from reckon API
                            return values["Message"];
                        }
                    }
                }

                //otherwise return the status code
                return response.StatusCode.ToString();

            }
            catch (Exception ex)
            {
                if (ex.Message == "One or more errors occurred.")
                {
                    return ex.InnerException.Message;
                }
                return ex.Message;
            }

        }


        public static async Task<string> RetrieveExecutedRAHCall(string requestId, string access_token)
        {
            try
            {

                var client = new HttpClient();
                //this is important gives the reckon server time to return SOMETHING....
                TimeSpan HttpTimeSpan = new TimeSpan(0, 0, 3, 0);
                client.Timeout = (HttpTimeSpan);

                // Request headers
                client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", ReckonApiHelper.SubscriptionKey);
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + access_token);



                var uri = "https://api.reckon.com/RAH/v2/{" + requestId + "}?";

                var response = await client.GetAsync(uri);

                //I don't know why but fuck me if this doesn't work to fix the Cannot Connect error....
                System.Threading.Thread.Sleep(2000);

                //Check the response status code and make sure that its sucessful
                var result = response.Content.ReadAsStringAsync().Result;

                if (!string.IsNullOrEmpty(result))
                {
                    if (result.StartsWith("[") || result.StartsWith("{"))
                    {
                        Dictionary<string, string> values = JsonConvert.DeserializeObject<Dictionary<string, string>>(result);
                        if (values.ContainsKey("RequestId"))
                        {
                            if (values["RetryLater"] == "True")
                            {
                                //means we need to retry again (later)
                                DialogResult rez = MessageBox.Show("Your Data is still processing on the Reckon Server. You can continue to wait or try again later. Do you want to continue to wait?", "Confirmation", MessageBoxButtons.YesNoCancel);
                                if (rez == DialogResult.Yes)
                                {
                                    //recursive call to this function to continue waiting
                                    return RetrieveExecutedRAHCall(requestId, access_token).Result;
                                }
                            }
                            else
                            {
                                //return the XML payload
                                return values["result"];
                            }
                        }
                        else if (values.ContainsKey("Message"))
                        {
                            //probably some error or something so return the message from reckon API
                            var retmsg = "";
                            values.TryGetValue("Message", out retmsg);
                            return retmsg;
                        }
                    }
                }

                //otherwise return the status code
                return response.StatusCode.ToString();

            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }

    }
}
