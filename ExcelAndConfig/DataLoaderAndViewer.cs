using System.Collections.Generic;
using System.Linq;
using LinqToExcel;
using System.Threading.Tasks;
using System.Configuration;
using System;

namespace ExcelAndConfig
{
    internal class DataLoaderAndViewer
    {
        private static string folderPath = ConfigurationManager.AppSettings["folder"];
        private static List<List<Row>> ExcelFile = new List<List<Row>>();
        private static List<Row> filteredData = new List<Row>();

        internal async static Task LoadDataAsync(string fileName ,string sheetName)
        {
            string path = folderPath+$@"\{fileName}.xlsx";
            ExcelQueryFactory connection = new ExcelQueryFactory(path);
            await Task.Run(() => ExcelFile.Add(connection.Worksheet(sheetName).Select(row => row).ToList()));
            connection.Dispose();
        }
        internal static void ViewAllData()
        {
            foreach ( List<Row> sheet in ExcelFile )
            {
                foreach( Row row in sheet )
                {
                    Console.WriteLine("Id: "+row["ID"]+"\tCategory: " + row["Category"]+ "\tProduct: " + row["Product"]);
                }
            }
        }
        private static void GetInputs(out string id,out string region,out string city,out string quantity,out string product,out string category)
        {
            string msg = "(Just press Enter to not filter by this option)";
            Console.WriteLine("Product ID: " + msg);
            id = Console.ReadLine();

            Console.WriteLine("Region: " + msg);
            region = Console.ReadLine();

            Console.WriteLine("City: " + msg);
            city = Console.ReadLine();

            Console.WriteLine("Quantity: " + msg);
            quantity = Console.ReadLine();

            Console.WriteLine("Category: " + msg);
            category = Console.ReadLine();

            Console.WriteLine("Product: " + msg);
            product = Console.ReadLine();
        }
        internal static void FilterData()
        {
            GetInputs(out string id, out string region, out string city, out string quantity, out string product, out string category);
            
            foreach(List<Row> sheet in ExcelFile)
            {
                List<Row> temp = sheet;
                if (id != null && id != string.Empty)
                {
                    temp = temp.Where(row => row["ID"].ToString() == id).ToList();
                }
                if (region != null && region != string.Empty)
                {
                    temp = temp.Where(row => row["Region"].ToString() == region).ToList();
                }
                if (city != null && city != string.Empty)
                {
                    temp = temp.Where(row => row["City"].ToString() == city).ToList();
                }
                if (quantity != null && quantity != string.Empty)
                {
                    temp = temp.Where(row => row["Qty"].ToString() == quantity).ToList();
                }
                if (category != null && category != string.Empty)
                {
                    temp = temp.Where(row => row["Category"].ToString() == category).ToList();
                }
                if (product != null && product != string.Empty)
                {
                    temp = temp.Where(row => row["Product"].ToString() == product).ToList();
                }
                foreach(Row row in temp) filteredData.Add(row);
            }
        }
        internal static void ViewFilteredData()
        {
            if (filteredData.Count() > 0 && filteredData != null)
            {
                Console.WriteLine("  {0}\t\t{1}\t   {2}\t{3}{4}{5}{6}", "ID".PadRight(5)
                                                                        , "Date".PadRight(15), "Region".PadRight(5)
                                                                        , "City".PadRight(10), "\tProduct".PadRight(20)
                                                                        , "Category".PadRight(15), "Quantity".PadRight(5));

                foreach (Row row in filteredData)
                {
                    Console.WriteLine(row["ID"].ToString().PadRight(10) + " " +
                            row["Date"].ToString().PadRight(25) + " " +
                            row["Region"].ToString().PadRight(10) + " " +
                            row["City"].ToString().PadRight(10) + " \t" +
                            row["Product"].ToString().PadRight(20) + " " +
                            row["Category"].ToString().PadRight(15) + " " +
                            row["Qty"].ToString().PadRight(5));
                }
                Console.WriteLine();
            }
            else Console.WriteLine("--------No Relevant Data Found--------\n");
            filteredData = new List<Row>();
        }
    }
}
