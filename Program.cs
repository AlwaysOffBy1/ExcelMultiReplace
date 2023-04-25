using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        string? oldVal = "";
        string? newVal = "";
        List<string> filePaths;

        while (oldVal == "" || oldVal is null)
        {
            Console.WriteLine("What is the value you are searching for?");
            oldVal = Console.ReadLine();
        }
        while (newVal == "" || newVal is null)
        {
            Console.WriteLine("What is the value you are replacing " + oldVal + " with?");
            newVal = Console.ReadLine();
        }

        try
        {
            filePaths = Directory.GetFiles(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "*.xl*", SearchOption.AllDirectories).ToList();

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            foreach(string s in filePaths)
            {

            }
        }
        catch(Exception e)
        {
            
        }


    }
}
