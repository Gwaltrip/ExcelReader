using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;


namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            var file = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase),"ExcelBook.xlsx");
            ExcelReader exr = new ExcelReader(new Uri(file).LocalPath);
            List<double> list = exr.Range(1);

            Console.WriteLine("Average: " + list.Average());
            Console.WriteLine("Sum: " + list.Sum());
            Console.ReadKey();
        }
    }
}
