using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

namespace AWSLoadBalancerLoadTester
{
    public class LoadTester
    {
        string URL = null;
        int max_reads = 10;
        int max_threads = 10;

        public LoadTester(string url, int maxReads, int maxThreads)
        {
            URL = url;
            max_reads = maxReads;
            max_threads = maxThreads;
        }

        public void StartLoadTest()
        {
            for (int iIndex = 0; iIndex < max_threads; iIndex++)
            {
                Thread newThread = new Thread(AddLoad);
                newThread.Start();
            }
        }

        private void AddLoad(object obj)
        {            
            var thread_id = Thread.CurrentThread.ManagedThreadId;
            for (int iIndex = 0; iIndex < max_reads; iIndex++)
            {
                Console.WriteLine(string.Format("Thread ID: {0}, Iteration: {1}, Output: {2}", thread_id.ToString(), iIndex, Get(URL)));
            }            
        }

        public string Get(string uri)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }
}
