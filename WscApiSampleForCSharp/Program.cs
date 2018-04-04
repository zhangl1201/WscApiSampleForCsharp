using System;

namespace WscApiSampleForCSharp
{
    class Program
    {
        public static void Main(string[] args)
        {
            WscProductInfoAgent wscAgent = new WscProductInfoAgent();
            wscAgent.CheckForSecurityCenterProducts();
            Console.ReadKey();
        }
    }
}
