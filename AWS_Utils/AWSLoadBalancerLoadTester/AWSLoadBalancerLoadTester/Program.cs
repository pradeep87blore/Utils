using System;

namespace AWSLoadBalancerLoadTester
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting Load Test");

            LoadTester loadTester = new LoadTester("http://alb-1646804891.us-east-1.elb.amazonaws.com/", 10, 20);
            loadTester.StartLoadTest();
        }
    }
}
