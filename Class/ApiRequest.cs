using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web;

namespace MigratorAzureDevops
{
    public class APIRequest
    {
        HttpClient client;
        //Constructor for initialising Client
        public APIRequest(string PAT)
        {
            client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(
                           new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                Convert.ToBase64String(
                    System.Text.ASCIIEncoding.ASCII.GetBytes(
                        string.Format("{0}:{1}", "", PAT))));
        }
        public string ApiRequest(string url, string method = "GET", string requestBody = null)
        {
            HttpRequestMessage Request;
            if (requestBody != null)
            {
                HttpContent content = new StringContent(requestBody, Encoding.UTF8, "application/json");
                Request = new HttpRequestMessage(new HttpMethod(method), url) { Content = content };
            }
            else
            {
                Request = new HttpRequestMessage(new HttpMethod(method), url);
            }
            using (HttpResponseMessage response = client.SendAsync(Request).Result)
            {
                if (response.IsSuccessStatusCode)
                    return response.Content.ReadAsStringAsync().Result;
                else
                    return null;
            }
        }
    }
}