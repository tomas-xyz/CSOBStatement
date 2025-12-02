using System.Diagnostics;
using System.Text.Json;

namespace tomxyz.csob;

class Program
{
    private static readonly string ConfigFileName = "config.json";
    public static async Task ProcessStatementAsync(string xmlStatementPath, IEnumerable<string> additionalPaths, Configuration configuration)
    {
        var stopwatch = new Stopwatch();

        stopwatch.Start();

        var statement = new Statement(xmlStatementPath);

        var movements = statement.Movements?.ToList() ?? new List<Movement>();
        foreach (var path in additionalPaths)
        {// add movements from additional accounts
            var add = new Statement(path);
            movements.AddRange(add.Movements);
        }

        statement.Movements = movements;

        var gs = new GsheetCSOB(configuration.SheetId);
        await gs.AutenticateAsync("key.json");

        var ownAccounts = await gs.ReadStringRangeAsync(configuration.Accounts);
        var categories = await gs.ReadStringRangeAsync(configuration.Categories);
        var rules = await gs.ReadRulesAsync(configuration.Rules);

        var orderedCategorized = statement.Movements
            .ToLookup(x => x.GetCategory(rules))
            .OrderBy(x => x.Key, new CategoriesComparer(categories));

        var (newTab, sheetId) = await gs.CreateStatementTab(statement.DateFrom);

        var startRow = await gs.WriteSummaryAsync(newTab, sheetId, statement);
        await gs.WriteMovements(startRow + 1, newTab, sheetId, orderedCategorized, configuration.Categories, ownAccounts);
        await gs.WriteStatisticsAsync(newTab, sheetId, configuration.Categories, categories);

        stopwatch.Stop();
        Console.WriteLine($"Výpis '{xmlStatementPath} 'zpracován do listu '{newTab}' v čase: {stopwatch.Elapsed}");
    }

    public static int Main(string[] args)
    {
        var filePath = string.Empty;
        try
        {
            if (args.Length < 1)
                throw new Exception($"Program očekává jeden parametr s cestou k xml souboru výpisu ČSOB");

            filePath = args[0];
            var additionalPaths = args[1..];

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

            ProcessStatementAsync(filePath, additionalPaths, configuration)
                .GetAwaiter()
                .GetResult();

            return 0;
        }
        catch (Exception e)
        {
            Console.WriteLine($"Zpracování výpisu '{filePath}' selhalo s chybou: \r\n {e}");
            return 1;
        }
    }
}
