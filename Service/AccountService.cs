using distribution_copy.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web;


namespace MigratorAzureDevops.Service
{
    public class AccountService
    {
        public string PAT;
        public AccountService()
        {
            this.PAT = HttpContext.Current.Session["PAT"] == null ? "" : HttpContext.Current.Session["PAT"].ToString();
        }
        public string GenerateRequestPostData(string appSecret, string authCode, string callbackUrl)
        {
            try
            {
                return String.Format("client_assertion_type=urn:ietf:params:oauth:client-assertion-type:jwt-bearer&client_assertion={0}&grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion={1}&redirect_uri={2}",
                            HttpUtility.UrlEncode(appSecret),
                            HttpUtility.UrlEncode(authCode),
                            callbackUrl
                     );
            }
            catch (Exception)
            {
                //ViewBag.ErrorMessage = ex.Message;
            }
            return string.Empty;
        }

        /// <summary>
        /// Generate Access Token
        /// </summary>
        /// <param name="body"></param>
        /// <returns></returns>
        public AccessDetails GetAccessToken(string body)
        {
            try
            {
                string baseAddress = System.Configuration.ConfigurationManager.AppSettings["BaseAddress"];
                var client = new HttpClient
                {
                    BaseAddress = new Uri(baseAddress)
                };

                var request = new HttpRequestMessage(HttpMethod.Post, "/oauth2/token");

                var requestContent = body;
                request.Content = new StringContent(requestContent, Encoding.UTF8, "application/x-www-form-urlencoded");

                var response = client.SendAsync(request).Result;
                if (response.IsSuccessStatusCode)
                {
                    string result = response.Content.ReadAsStringAsync().Result;
                    AccessDetails details = Newtonsoft.Json.JsonConvert.DeserializeObject<AccessDetails>(result);
                    return details;
                }
            }
            catch (Exception)
            {
                //ViewBag.ErrorMessage = ex.Message;
            }
            return new AccessDetails();
        }
        public AccountsResponse.AccountList GetAccounts(string memberID, AccessDetails details)
        {
            AccountsResponse.AccountList accounts = new AccountsResponse.AccountList();
            var client = new HttpClient();
            string baseAddress = System.Configuration.ConfigurationManager.AppSettings["BaseAddress"];

            string requestContent = baseAddress + "/_apis/Accounts?memberId=" + memberID + "&api-version=4.1";
            var request = new HttpRequestMessage(HttpMethod.Get, requestContent);
            request.Headers.Add("Authorization", "Bearer " + details.access_token);
            try
            {
                var response = client.SendAsync(request).Result;
                if (response.IsSuccessStatusCode && response.StatusCode == HttpStatusCode.OK)
                {
                    string result = response.Content.ReadAsStringAsync().Result;
                    accounts = JsonConvert.DeserializeObject<distribution_copy.Models.AccountsResponse.AccountList>(result);
                    return accounts;
                }
                else
                {
                    var errorMessage = response.Content.ReadAsStringAsync();
                }
            }
            catch (Exception)
            {
            }
            return accounts;
        }
        public ProfileDetails GetProfile(AccessDetails accessDetails)
        {
            ProfileDetails profile = new ProfileDetails();
            using (var client = new HttpClient())
            {
                try
                {
                    string baseAddress = System.Configuration.ConfigurationManager.AppSettings["BaseAddress"];
                    client.BaseAddress = new Uri(baseAddress);
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessDetails.access_token);
                    HttpResponseMessage response = client.GetAsync("_apis/profile/profiles/me?api-version=4.1").Result;
                    if (response.IsSuccessStatusCode && response.StatusCode == HttpStatusCode.OK)
                    {
                        string result = response.Content.ReadAsStringAsync().Result;
                        profile = JsonConvert.DeserializeObject<ProfileDetails>(result);
                        return profile;
                    }
                    else
                    {
                        var errorMessage = response.Content.ReadAsStringAsync();
                    }
                }
                catch (Exception)
                {
                }
                return profile;
            }
        }

        public T GetApi<T>(string url, string method = "GET", string requestBody = null)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.ASCII.GetBytes(string.Format("{0}:{1}", "", PAT))));
            HttpRequestMessage Request;
            if (requestBody != null)
            {
                HttpContent content = new StringContent(requestBody, Encoding.UTF8, "application/json");
                Request = new HttpRequestMessage(new HttpMethod(method), url) { Content = content };
            }
            else
                Request = new HttpRequestMessage(new HttpMethod(method), url);

            using (HttpResponseMessage response = client.SendAsync(Request).Result)
            {
                if (response.IsSuccessStatusCode)
                {
                    var responseBody = response.Content.ReadAsStringAsync().Result;
                    return JsonConvert.DeserializeObject<T>(responseBody);
                }
                else
                    return default;
            }
        }
    }
}