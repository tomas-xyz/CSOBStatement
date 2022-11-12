using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace tomxyz.csob;

class Program
{
    public static async Task ProcessStatementAsync(string xmlStatementPath)
    {
        var statement = new Statement(xmlStatementPath);
        
        

        var categories = await gs.ReadCategoriesAsync();
        var rules = await gs.ReadRulesAsync();

        var categorized = statement.Movements.ToLookup(x => x.GetCategory(rules));
        var (newTab, sheetId) = ("a", 468917263);// await gs.CreateStatementTab(statement.DateFrom);        
        await gs.WriteMovements(newTab, sheetId, categorized, categories);
    }

    public static int Main(string[] args)
    {
        try
        {
            if (args.Length == 0)
                throw new Exception("Pass path to xml tatement as agrument");

            ProcessStatementAsync(args[0])
                .GetAwaiter()
                .GetResult();

            return 0;
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
            return 1;
        }
    }
}
