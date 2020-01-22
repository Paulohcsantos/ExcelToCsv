using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToCsv
{
    class Program
    {
        static ExcelConvert excelConvert = new ExcelConvert();

        static void Main(string[] args)
        {
            // File In (xls, xlsx) , File Out (Csv)
            excelConvert.Convert(@"f:\extrato_ingrid.xls", @"f:\extrato_ingrid.csv");
        }
    }
}
