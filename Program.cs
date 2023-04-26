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
            var exeLoc = System.AppContext.BaseDirectory;
            filePaths = Directory.GetFiles(exeLoc,"*.xl*",SearchOption.AllDirectories).ToList().OrderBy(x => x).ToList();

        }
        catch(Exception e)
        {
            Console.WriteLine(e.ToString());
            Console.WriteLine("The program will now exit");
            Console.ReadLine();
            System.Environment.Exit(-1);
        }
        Console.WriteLine("Found " + filePaths.Count + " excel files");
        Microsoft.Office.Interop.Excel.Application? xlApp = null;
        Workbook? wb = null;


        //Excel files found, try to run loop to replace text from user params
        //TODO what if file is currently open?
        try 
        { 
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            foreach(string s in filePaths)
            {
                Console.WriteLine("Opening file " + s);
                while(wb is null) { TryOpenWorkbook(xlApp, s); }
                int len = wb.Worksheets.Count;
                foreach(Worksheet ws in wb.Worksheets)
                {
                    Console.WriteLine("  Opening sheet " + ws.Name);
                    Microsoft.Office.Interop.Excel.Range r = ws.UsedRange;
                    bool success = (bool)r.Replace2(oldVal, newVal, XlLookAt.xlWhole, XlSearchOrder.xlByRows, false, false, false, false, false);
                }
                wb.Close(true);
            }
        }
        catch(Exception e)
        {
            Console.WriteLine(e.ToString());
            Console.WriteLine("The program will now exit");
        }
        finally
        {
            if(xlApp is not null) xlApp.Quit();
            xlApp = null;
        }
    }
    public static Workbook? TryOpenWorkbook(Application xl,string path)
    {
        using(CancellationTokenSource cts = new CancellationTokenSource(5000))
        {
            cts.Token.Register(() => 
            {
                Console.WriteLine($"Workbook {path} cannot be opened. It may currently be in use. Please close the file then press enter");
                Console.ReadLine();

            });
            while (!cts.IsCancellationRequested)
            {
                return xl.Workbooks.Open(path);
            }
        }
        return null;
    }
}


