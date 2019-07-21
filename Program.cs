using System;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            string filepath = AppDomain.CurrentDomain.BaseDirectory + @"Template\Template311.xls";
            string filepath2 = AppDomain.CurrentDomain.BaseDirectory + @"Template\Template311x.xlsx";

            var pdfhelp = new ToPdfHelper("xls");

            pdfhelp.XlsWpsToPdf(filepath, "Template311.xls");

            Console.ReadKey();
        }
    }
}
