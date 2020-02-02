using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ReckonDesktop.Model;

namespace ReckonDesktop.Autofile
{
    public class AutoFile
    {
        public AutoFile(string url, string username, string password)
        {
            Url = url;
            Username = username;
            Password = password;
        }

        public static string Url { get; set; }

        public static string Username { get; set; }

        public static string Password { get; set; }

        public static async Task<string> GetApiToken()
        {
            using (var client = new HttpClient())
            {
                //setup client 
                client.BaseAddress = new Uri(Url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));


                //setup login data 
                var formContent = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("grant_type", "password"),
                    new KeyValuePair<string, string>("username", Username),
                    new KeyValuePair<string, string>("password", Password),
                });


                //send request 
                HttpResponseMessage responseMessage = await client.PostAsync("/Token", formContent);


                //get access token from response body 
                var responseJson = await responseMessage.Content.ReadAsStringAsync();
                var jObject = JObject.Parse(responseJson);
                return jObject.GetValue("access_token").ToString();
            }
        }

        public async Task<string> GetBills(int accountingSystem = 9)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback +=
                    (sender, cert, chain, sslPolicyErrors) => true;
                var token = await GetApiToken();
                //setup client 
                client.BaseAddress = new Uri(Url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                //var formContent = new FormUrlEncodedContent(new[]
                //{
                //    //new KeyValuePair<string, string>("grant_type", "password"),
                //    //new KeyValuePair<string, string>("username", userName),
                //    new KeyValuePair<string, string>("", ""),
                //});

                var apiPath = "api/v1/AccountingEntries/" + accountingSystem.ToString();
                HttpResponseMessage postResponse = await client.GetAsync(Url.TrimEnd('/') + "/" + apiPath);

//                postResponse.EnsureSuccessStatusCode();
                var postResult = await postResponse.Content.ReadAsStringAsync();
                return postResult.ToString();
                ////send request 
                //HttpResponseMessage responseMessage = await client.PostAsync("/Token", formContent);


                ////get access token from response body 
                //var responseJson = await responseMessage.Content.ReadAsStringAsync();
                //var jObject = JObject.Parse(responseJson);
                //return jObject.GetValue("access_token").ToString();
            }
        }

        public async Task<string> UpdateBill(int id, string token, int accountingSystem = 9)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback +=
                    (sender, cert, chain, sslPolicyErrors) => true;
                //                var token = await GetApiToken();
                //setup client 
                client.BaseAddress = new Uri(Url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                var apiPath = "api/v1/AccountingEntries/" + id.ToString();
                var param = "";
                HttpContent contentPost = new StringContent(param, Encoding.UTF8, "application/json");
                var response = await client.PostAsync(apiPath, contentPost).ConfigureAwait(false);
                return response.StatusCode.ToString();
                //HttpResponseMessage postResponse = await client.GetAsync(Url.TrimEnd('/') + "/" + apiPath);

                ////                postResponse.EnsureSuccessStatusCode();
                //var postResult = await postResponse.Content.ReadAsStringAsync();
                //return postResult.ToString();
                //////send request 
                ////HttpResponseMessage responseMessage = await client.PostAsync("/Token", formContent);


                //////get access token from response body 
                ////var responseJson = await responseMessage.Content.ReadAsStringAsync();
                ////var jObject = JObject.Parse(responseJson);
                ////return jObject.GetValue("access_token").ToString();
            }
        }

        public string SendContacts(AccountingContact contact, string token, int accountingSystem = 9)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback +=
                    (sender, cert, chain, sslPolicyErrors) => true;
//                var token = await GetApiToken();
                //setup client 
                client.BaseAddress = new Uri(Url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                var apiPath = "api/v1/AccountingContacts";
                var param = Newtonsoft.Json.JsonConvert.SerializeObject(contact);
                HttpContent contentPost = new StringContent(param, Encoding.UTF8, "application/json");
                var response = client.PostAsync(apiPath, contentPost).Result;
                return response.StatusCode.ToString();
                //HttpResponseMessage postResponse = await client.GetAsync(Url.TrimEnd('/') + "/" + apiPath);

                ////                postResponse.EnsureSuccessStatusCode();
                //var postResult = await postResponse.Content.ReadAsStringAsync();
                //return postResult.ToString();
                //////send request 
                ////HttpResponseMessage responseMessage = await client.PostAsync("/Token", formContent);


                //////get access token from response body 
                ////var responseJson = await responseMessage.Content.ReadAsStringAsync();
                ////var jObject = JObject.Parse(responseJson);
                ////return jObject.GetValue("access_token").ToString();
            }
        }

        public string SendAccounts(AccountingAccount contact, string token, int accountingSystem = 9)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback +=
                    (sender, cert, chain, sslPolicyErrors) => true;
                //                var token = await GetApiToken();
                //setup client 
                client.BaseAddress = new Uri(Url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                var apiPath = "api/v1/AccountingAccounts";
                var param = Newtonsoft.Json.JsonConvert.SerializeObject(contact);
                HttpContent contentPost = new StringContent(param, Encoding.UTF8, "application/json");
                var response = client.PostAsync(apiPath, contentPost).Result;
                return response.StatusCode.ToString();
                //HttpResponseMessage postResponse = await client.GetAsync(Url.TrimEnd('/') + "/" + apiPath);

                ////                postResponse.EnsureSuccessStatusCode();
                //var postResult = await postResponse.Content.ReadAsStringAsync();
                //return postResult.ToString();
                //////send request 
                ////HttpResponseMessage responseMessage = await client.PostAsync("/Token", formContent);


                //////get access token from response body 
                ////var responseJson = await responseMessage.Content.ReadAsStringAsync();
                ////var jObject = JObject.Parse(responseJson);
                ////return jObject.GetValue("access_token").ToString();
            }
        }

        public string SendItemTypes(AccountingItemType contact, string token, int accountingSystem = 9)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback +=
                    (sender, cert, chain, sslPolicyErrors) => true;
                //                var token = await GetApiToken();
                //setup client 
                client.BaseAddress = new Uri(Url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                var apiPath = "api/v1/AccountingItemTypes";
                var param = Newtonsoft.Json.JsonConvert.SerializeObject(contact);
                HttpContent contentPost = new StringContent(param, Encoding.UTF8, "application/json");
                var response = client.PostAsync(apiPath, contentPost).Result;
                return response.StatusCode.ToString();
            }
        }

        public string SendJobs(AccountingJob contact, string token, int accountingSystem = 9)
        {
            using (var client = new HttpClient())
            {
                ServicePointManager.ServerCertificateValidationCallback +=
                    (sender, cert, chain, sslPolicyErrors) => true;
                //                var token = await GetApiToken();
                //setup client 
                client.BaseAddress = new Uri(Url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                var apiPath = "api/v1/AccountingJobs";
                var param = Newtonsoft.Json.JsonConvert.SerializeObject(contact);
                HttpContent contentPost = new StringContent(param, Encoding.UTF8, "application/json");
                var response = client.PostAsync(apiPath, contentPost).Result;
                return response.StatusCode.ToString();
                //HttpResponseMessage postResponse = await client.GetAsync(Url.TrimEnd('/') + "/" + apiPath);

                ////                postResponse.EnsureSuccessStatusCode();
                //var postResult = await postResponse.Content.ReadAsStringAsync();
                //return postResult.ToString();
                //////send request 
                ////HttpResponseMessage responseMessage = await client.PostAsync("/Token", formContent);


                //////get access token from response body 
                ////var responseJson = await responseMessage.Content.ReadAsStringAsync();
                ////var jObject = JObject.Parse(responseJson);
                ////return jObject.GetValue("access_token").ToString();
            }
        }

        public async Task<string> UploadFile(string filePath, List<KeyValuePair<string, string>> keyValues,
            string userName, string password, string apiBaseUri)
        {
            try
            {
                HttpClient client = new HttpClient();
                ServicePointManager.ServerCertificateValidationCallback +=
                    (sender, cert, chain, sslPolicyErrors) => true;
                var token = await GetApiToken();
                //setup client 
                client.BaseAddress = new Uri(apiBaseUri);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("multipart/form-data"));
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);


                // Issue MIME multipart POST request with a MIME multipart message containing a single
                // body part with StringContent.
                //                StringContent content = new StringContent("Hello World", Encoding.UTF8, "text/plain");
                MultipartFormDataContent formData = new MultipartFormDataContent();

                var path = @filePath;
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                var memStream = new MemoryStream();
                memStream.SetLength(fs.Length);
                fs.Read(memStream.GetBuffer(), 0, (int) fs.Length);
                formData.Add(new StreamContent(memStream), "file", Path.GetFileName(filePath));
                foreach (var keyValue in keyValues)
                {
                    formData.Add(new StringContent(keyValue.Value, Encoding.UTF8), keyValue.Key);
                }
                //                formData.Add(new StringContent("7", Encoding.UTF8), "ProjectId");
                //                //            formData.Add(new StringContent("201515", Encoding.UTF8), "PO Number");
                //                formData.Add(new StringContent("Cory", Encoding.UTF8), "First name");
                //                formData.Add(new StringContent("Wilson", Encoding.UTF8), "Surname");

                //                Console.WriteLine("Uploading data to store...");
                var apiPath = "api/v1/Document";
                HttpResponseMessage postResponse = await client.PostAsync(apiBaseUri + apiPath, formData);

                postResponse.EnsureSuccessStatusCode();
                var postResult = await postResponse.Content.ReadAsStringAsync();
                return postResult.ToString();

                // Issue GET request to get the content back from the store
                //            Console.WriteLine("Retrieving data from store: {0}", location);
                //            HttpResponseMessage getResponse = await client.GetAsync(location);

                //            getResponse.EnsureSuccessStatusCode();
                //            string result = await getResponse.Content.ReadAsStringAsync();
                //                Console.WriteLine("Received response: {0}", location);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public async Task<byte[]> SendData(NameValueCollection values, string userName, string password,
            string apiBaseUri)
        {
            var token = await GetApiToken();
            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
            string apiPath = "api/v1/Document";
            var request = WebRequest.Create(Url + apiPath);
            request.Method = "POST";
            var boundary = "---------------------------" +
                           DateTime.Now.Ticks.ToString("x", NumberFormatInfo.InvariantInfo);
            request.ContentType = "multipart/form-data; boundary=" + boundary;
            boundary = "--" + boundary;
            request.Headers.Add("Authorization", "Bearer " + token);
            using (var requestStream = request.GetRequestStream())
            {
                // Write the values
                foreach (string name in values.Keys)
                {
                    var buffer = Encoding.ASCII.GetBytes(boundary + Environment.NewLine);
                    requestStream.Write(buffer, 0, buffer.Length);
                    buffer =
                        Encoding.ASCII.GetBytes(string.Format("Content-Disposition: form-data; name=\"{0}\"{1}{1}", name,
                            Environment.NewLine));
                    requestStream.Write(buffer, 0, buffer.Length);
                    buffer = Encoding.UTF8.GetBytes(values[name] + Environment.NewLine);
                    requestStream.Write(buffer, 0, buffer.Length);
                }

                var boundaryBuffer = Encoding.ASCII.GetBytes(boundary + "--");
                requestStream.Write(boundaryBuffer, 0, boundaryBuffer.Length);
            }

            using (var response = request.GetResponse())
            using (var responseStream = response.GetResponseStream())
            using (var stream = new MemoryStream())
            {
                responseStream?.CopyTo(stream);
                return stream.ToArray();
            }
        }
    }
}