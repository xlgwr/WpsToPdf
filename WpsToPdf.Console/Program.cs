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

            var pdfhelp = new ToPdfHelper(filepath);
            var pdfhelpx = new ToPdfHelper(filepath);

            var filename = pdfhelp.SavePdf("Template311");
            var filename2 = pdfhelpx.SavePdf("Template311x");

            Console.WriteLine("生成pdf成功!" + filename);
            Console.WriteLine("\n生成pdf成功!" + filename2);

            Console.ReadKey();
        }
    }
}
