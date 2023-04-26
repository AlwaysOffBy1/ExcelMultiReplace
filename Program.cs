using Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        string? oldVal = null;
        string? newVal = null;
        List<string>? filePaths = new List<string>();


        //While loops to ensure user entered params aren't null
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

        //Create xlApp and workbook obj
        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        Workbook wb;


        //Excel files found, try to run loop to replace text from user params
        foreach(string s in filePaths)
        {
            Console.WriteLine("Opening file " + s);
            wb = TryOpenWorkbook(xlApp, s); 
            int len = wb.Worksheets.Count;
            foreach(Worksheet ws in wb.Worksheets)
            {
                Console.WriteLine("  Opening sheet " + ws.Name);
                Microsoft.Office.Interop.Excel.Range r = ws.UsedRange;
                bool success = (bool)r.Replace2(oldVal, newVal, XlLookAt.xlWhole, XlSearchOrder.xlByRows, false, false, false, false, false);
            }
            wb.Close(true);
        }
        if (xlApp is not null) xlApp.Quit();
    }
    public static Workbook TryOpenWorkbook(Application xl,string path)
    {
        Exception? e = null;
        Workbook? workbook = null;
        while (workbook is null)
        {
            using (CancellationTokenSource cts = new CancellationTokenSource(5000))
            {
                cts.Token.Register(() =>
                {
                    //When time is up
                });
                //Long task to try
                //Task is within a while loop to ensure workbook is closed before proceeding
                while (!cts.IsCancellationRequested && e is null)
                {
                    try
                    {
                        return xl.Workbooks.Open(path);
                    }
                    catch (Exception ex)
                    {
                        e = ex;
                    }
                }
            }
            if (e is not null)
            {
                Console.WriteLine($"Workbook {path} cannot be opened. It may currently be in use. Please close the file then press enter");
                e = null;
                Console.ReadLine();
            }
        }
        return workbook;
    }
}


