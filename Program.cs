using System.Text.Json;

namespace tomxyz.csob;

class Program
{
    private static string ConfigFileName = "config.json";
    public static async Task ProcessStatementAsync(string xmlStatementPath, Configuration configuration)
    {
        var statement = new Statement(xmlStatementPath);
        
        var gs = new GsheetCSOB(configuration.SheetId);
        await gs.AutenticateAsync("key.json");

        var categories = await gs.ReadCategoriesAsync(configuration.Categories);
        var rules = await gs.ReadRulesAsync(configuration.Rules);

        var categorized = statement.Movements.ToLookup(x => x.GetCategory(rules));
        var (newTab, sheetId) = await gs.CreateStatementTab(statement.DateFrom);

        var startRow = await gs.WriteSummaryAsync(newTab, sheetId, statement);
        await gs.WriteMovements(startRow + 1, newTab, sheetId, categorized, configuration.Categories);
    }

    public static int Main(string[] args)
    {
        try
        {
            var configuration = new Configuration();
            try
            {
                using var file = File.OpenRead(ConfigFileName);
                configuration = JsonSerializer.Deserialize<Configuration>(file);
                if (configuration == null)
                    throw new Exception($"Konfigurační soubor '{ConfigFileName}' se nebylo možné zpracovat");
            }
            catch (Exception e)
            {
                throw new Exception($"Konfigurační soubor '{ConfigFileName}' neexistuje nebo se jej nepodařilo zpracovat", e);
            }

            if (args.Length == 0)
                throw new Exception("Pass path to xml tatement as agrument");

            ProcessStatementAsync(args[0], configuration)
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
