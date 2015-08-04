using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hans.ExcelToJSON
{
    class Program
    {
        static void Main(string[] args)
        {
            //mongoimport --db DataStore --collection users --file File_2015-04-21-142804.json
            var path = ConfigurationManager.AppSettings["Path"].ToString(); ;
            var sheet = ConfigurationManager.AppSettings["Sheet"].ToString(); ;
            var fileName = ConfigurationManager.AppSettings["FileName"].ToString();

            var excel = new ExcelQueryFactory(path);
            var columns = excel.GetColumnNames(sheet);
            var list = excel.Worksheet(sheet).ToList();

            var headers = new List<string>();

            // get header
            foreach (var c in columns)
            {
                if (!string.IsNullOrEmpty(c.ToString()))
                {
                    var header = c.ToString().Trim().ToLower()
                        .Replace(" ", "")
                        .Replace("/", "")
                        .Replace("(", "")
                        .Replace(")", "");

                    headers.Add(header);
                }
            }

            Console.WriteLine();
            Console.WriteLine("[Convert Excel to JSON]");
            Console.WriteLine("=======================");
            Console.WriteLine("Path         : {0}", path);
            Console.WriteLine("Sheet        : {0}", sheet);
            Console.WriteLine("File Name    : {0}", fileName);
            Console.WriteLine();
            Console.WriteLine("Generating...");
            Console.WriteLine();

            var outputFileName = string.Format("{0}_{1}.json", fileName, DateTime.Now.ToString("yyyy-MM-dd-HHmmss"));
            var outFile = System.IO.File.CreateText(outputFileName);
            
            foreach (var item in list)
            {
                // generate json
                outFile.WriteLine("{");

                for (int i = 0; i < headers.Count; i++)
                {
                    if (i != (headers.Count - 1))
                    {
                        outFile.WriteLine("    \"{0}\": \"{1}\",", headers[i], item[i].ToString()
                            .Replace("\n", " ")
                            .Replace("\r", " ")
                            .Trim());
                    }
                    else
                    {
                        outFile.WriteLine("    \"{0}\": \"{1}\"", headers[i], item[i].ToString()
                            .Replace("\n", " ")
                            .Replace("\r", " ")
                            .Trim());
                    }
                }

                outFile.WriteLine("}");
            }

            outFile.Close();

            Console.WriteLine(string.Format("{0} created", outputFileName));
            Console.WriteLine("Press any key to continue...");
            Console.Read();
        }
    }
}
