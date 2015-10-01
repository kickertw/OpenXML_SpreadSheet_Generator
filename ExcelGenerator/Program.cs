using Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = new DataSet();
            var sheet1 = new DataTable();
            var sheet2 = new DataTable();
            List<string> sheetNames = new List<string>() { "First Sheet", "Second Sheet" };

            //Sheet 1 Headers
            sheet1.Columns.Add("Test Header For Strings", typeof(string));
            sheet1.Columns.Add("Test Header For Numbers", typeof(int));
            sheet1.Columns.Add("Test Header For Decimals", typeof(double));
            sheet1.Columns.Add("Test Header For HyperLinks", typeof(string));
            
            //Sheet 1 Data
            sheet1.Rows.Add("Test Val 1", 1, 1.50, "http://www.google.com/search?q=Ben+Carson");
            sheet1.Rows.Add("Test Val 2", 160, 34.50, "http://www.google.com/search?q=Donald+Trump");
            sheet1.Rows.Add("Test Val 3", 540, 9.99, "http://www.google.com/search?q=Hillary+Clinton");

            sheet2.Columns.Add("Test Header For Strings", typeof(string));
            sheet2.Columns.Add("Test Header For Numbers", typeof(int));
            sheet2.Columns.Add("Test Header For Decimals", typeof(double));
            sheet2.Columns.Add("Test Header For HyperLinks", typeof(string));

            //Sheet 1 Data
            sheet2.Rows.Add("Ichigo", 1, 1.50, "http://www.google.com/search?q=Bankai");
            sheet2.Rows.Add("Goku", 160, 34.50, "http://www.google.com/search?q=Dattebayo");

            data.Tables.Add(sheet1);
            data.Tables.Add(sheet2);

            var wbData = new WorkBookData(data, sheetNames);

            var xlGen = new Utils.ExcelGenerator();
            xlGen.CreateSpreadSheet(@"C:\Users\twong\Documents\xlgentest\output.xlsx", wbData);
        }
    }
}
