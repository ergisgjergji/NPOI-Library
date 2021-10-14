using Examples.Models;
using Examples.Tests;
using Npoi_Library.Excel;
using Npoi_Library.Excel.Configurations;
using Npoi_Library.Excel.Styling;
using Npoi_Library.Excel.XlsManager;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            XlsxTests.Test1();
            //Test2();
            //Test3();
            //Test_Read();

            Console.WriteLine("Tests completed successfully!");
            Console.ReadLine();
        }

        
    }
}
