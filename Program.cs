using System;
using System.IO;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        string? oldVal = null;
        string? newVal = null;
        List<string>? filePaths = new List<string>();


        //While loops to ensure params aren't null
        while (oldVal == "" || oldVal is null)
        {
            Console.WriteLine("What is the value you are searching for?");
            oldVal = Console.ReadLine();
        }
        while (newVal is null)
        {
            Console.WriteLine("What is the value you are replacing " + oldVal + " with?");
            newVal = Console.ReadLine();
        }

        //Try to access excel files
        try
        {
            filePaths = Directory.GetFiles(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "*.xl*", SearchOption.AllDirectories).ToList();
        }
        catch(Exception e)
        {
            Console.WriteLine(e.ToString());
            Console.WriteLine("The program will now exit");
            Console.ReadLine();
            System.Environment.Exit(-1);
        }
        Console.WriteLine("Found " + filePaths.Count + " excel files");
        Microsoft.Office.Interop.Excel.Application xlApp = null;

        //Excel files found, try to run loop to replace text from user params
        try 
        { 
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            foreach(string s in filePaths)
            {
                Console.WriteLine("Opening file " + s);
                Workbook wb = xlApp.Workbooks.Open(s);
                int len = wb.Worksheets.Count;
                foreach(Worksheet ws in wb.Worksheets)
                {
                    Console.WriteLine("  Opening sheet " + ws.Name);
                    Microsoft.Office.Interop.Excel.Range r = ws.UsedRange;
                    bool success = (bool)r.Replace2(oldVal, newVal, XlLookAt.xlWhole, XlSearchOrder.xlByRows, false, false, false, false, false);
                }
                wb.Close(true);
            }
            xlApp.Quit();
            xlApp = null;
        }
        catch(Exception e)
        {
            Console.WriteLine(e.ToString());
            Console.WriteLine("The program will now exit");
            xlApp = null;
        }
    }
}
