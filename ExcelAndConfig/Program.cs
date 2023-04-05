using System.Collections.Generic;
using System.Threading.Tasks;
using System.Configuration;
using System;

namespace ExcelAndConfig
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string[] fileNames = ConfigurationManager.AppSettings["files"].Split(',');
            string[] sheetNames = ConfigurationManager.AppSettings["sheets"].Split(',');
            List<Task> tasks = new List<Task>();
            foreach (string file in fileNames)
            {
                foreach (string sheet in sheetNames)
                {
                    tasks.Add(DataLoaderAndViewer.LoadDataAsync(file, sheet));
                }
            }
            await Task.WhenAll(tasks);
            //DataLoaderAndViewer.ViewAllData();
            while (true)
            {
                Console.WriteLine("Enter 'q' to exit or press any other key to continue...");
                string quit = Console.ReadLine();
                if (quit.ToUpper() != "Q")
                {
                    DataLoaderAndViewer.FilterData();
                    DataLoaderAndViewer.ViewFilteredData();
                }
                else break;
            }
        }
    }
}
