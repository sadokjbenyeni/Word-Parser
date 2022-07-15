using Newtonsoft.Json;
using Services.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace ConsoleApp1
{
    class Program
    {

        //**************************Methode1**********************
        //    static void Main(string[] args)
        //    {
        //        PostRequest("http://httpbin.org/post");
        //        Console.ReadKey();
        //    }
        //    async static void PostRequest(string url)
        //    {
        //        IEnumerable<KeyValuePair<string, string>> queries = new List<KeyValuePair<string, string>>()
        //        {
        //            new KeyValuePair<string, string>("query1","Ahmed"),
        //            new KeyValuePair<string, string>("query2", "Houssem")

        //        };
        //        HttpContent q = new FormUrlEncodedContent(queries);
        //        using (HttpClient client = new HttpClient())
        //        {
        //            using (HttpResponseMessage response = await client.PostAsync(url, q))
        //            {
        //                using (HttpContent content = response.Content)
        //                {
        //                    string mycontent = await content.ReadAsStringAsync();
        //                    HttpContentHeaders headers = content.Headers;

        //                    Console.WriteLine(mycontent);

        //                }
        //            }
        //        }
        //    }

        //***********************Methode2****************************


        public static void Main()
        {

        }

        //public class Account
        //{
        //    public string Email { get; set; }
        //    public bool Active { get; set; }
        //    public DateTime CreatedDate { get; set; }
        //    public IList<string> Roles { get; set; }
        //}

        //***************************************Authentification Method*****************************************

        //public static void Main()
        //{
        //    Uri myUri = new Uri("http://httpbin.org/post");
        //    WebRequest myWebRequest = HttpWebRequest.Create(myUri);

        //    HttpWebRequest myHttpWebRequest = (HttpWebRequest)myWebRequest;

        //    NetworkCredential myNetworkCredential = new NetworkCredential("1i3tw-1543333616", "2018-11-27 15:46:56 UTC");

        //    CredentialCache myCredentialCache = new CredentialCache();
        //    myCredentialCache.Add(myUri, "Basic", myNetworkCredential);

        //    myHttpWebRequest.PreAuthenticate = true;
        //    myHttpWebRequest.Credentials = myCredentialCache;

        //    WebResponse myWebResponse = myWebRequest.GetResponse();

        //    Stream responseStream = myWebResponse.GetResponseStream();

        //    StreamReader myStreamReader = new StreamReader(responseStream, Encoding.Default);

        //    string pageContent = myStreamReader.ReadToEnd();

        //    responseStream.Close();

        //    myWebResponse.Close();

        //    Console.WriteLine(pageContent);

        public void HttpRequest(string projectId)
        {
            // Create a request using a URL that can receive a post. 
            string URL;
            URL = "http://emeagvaqc01.newaccess.ch/rest-api/service/api/v1/projects/" + projectId + "/tests";
            WebRequest request = WebRequest.Create("http://httpbin.org/post");
            // Set the Method property of the request to POST.
            request.Method = "POST";
            // Create POST data and convert it to a byte array.
            //***************
            request.ContentType = "application/json";
            request.ContentLength = 375;

            request.Headers["Authorization"] = "Basic YWRtaW5RQUNAbmV3YWNjZXNzLmNoOlFBX05XQTIwMTg=";
            request.Headers["Accept"] = "application/json";
            request.Headers["Connection"] = "keep-alive";




            //***************
            string json = "{'Email': 'james@example.com','Active': true,'CreatedDate':'2013-01-20T00:00:00Z','Roles': ['User','Admin']}";

            //Account account = JsonConvert.DeserializeObject<Account>(json);

            //string postData = "This is a test that posts this string to a Web server.";

            byte[] byteArray = Encoding.UTF8.GetBytes(json);
            // Set the ContentType property of the WebRequest.
            request.ContentType = "application/x-www-form-urlencoded";
            // Set the ContentLength property of the WebRequest.
            request.ContentLength = byteArray.Length;
            // Get the request stream.
            Stream dataStream = request.GetRequestStream();
            // Write the data to the request stream.
            dataStream.Write(byteArray, 0, byteArray.Length);
            // Close the Stream object.
            dataStream.Close();

            // Get the response.
            WebResponse response = request.GetResponse();
            // Display the status.
            Console.WriteLine(((HttpWebResponse)response).StatusDescription);
            // Get the stream containing content returned by the server.
            dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            // Display the content.
            Console.WriteLine(responseFromServer);

            // Clean up the streams.
            reader.Close();
            dataStream.Close();
            response.Close();
        }







    }


}
