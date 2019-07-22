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
            var pdfhelpx = new ToPdfHelper("xlsx");

            var filename = pdfhelp.XlsWpsToPdf(filepath, "Template311.xls");
            var filename2 = pdfhelpx.XlsWpsToPdf(filepath2, "Template311x.xlsx");

            Console.WriteLine("生成pdf成功!" + filename);

            Console.ReadKey();
        }
    }
}
