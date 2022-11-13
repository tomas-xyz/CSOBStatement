using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace tomxyz.csob;
internal class GsheetCSOB
{
    readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    private SheetsService? Service { get; set; }
    private GoogleCredential? Credentials { get; set; }
    private string GSheetId { get; }

    internal GsheetCSOB(string gsheetId)
    {
        GSheetId = gsheetId;
    }
        
    public async Task AutenticateAsync(string credentialFilePath, CancellationToken ct = default)
    {
        var credentials = await GoogleCredential.FromFileAsync(credentialFilePath, ct);
        if (credentials == null)
            throw new Exception("Unable to get credentials");

        Service = new SheetsService(new BaseClientService.Initializer()
         {
             HttpClientInitializer = credentials,
             ApplicationName = "CSOBStatement"
         });
    }

    public async Task<IEnumerable<string>> ReadCategoriesAsync(string categoriesRange, CancellationToken ct = default)
    {
        if (Service == null)
            throw new Exception("Google sheet service has not been authenticated");

        var valuesService = Service.Spreadsheets.Values;
        var request = valuesService.Get(GSheetId, categoriesRange);
        var response = await request.ExecuteAsync(ct);
        
        return response.Values.SelectMany(x => x).Select(o => (string)o);
    }

    public async Task<IEnumerable<Rule>> ReadRulesAsync(string rulesRange, CancellationToken ct = default)
    {
        if (Service == null)
            throw new Exception("Google sheet service has not been authenticated");

        var valuesService = Service.Spreadsheets.Values;
        var request = valuesService.Get(GSheetId, rulesRange);
        var response = await request.ExecuteAsync(ct);

        return response.Values
            .Skip(1)
            .Select(x => new Rule((string)x[0], (string)x[1], (string)x[2], (string)x[3]));
    }

    protected async Task PerformRequestAsync(Request request)
    {
        var update = new BatchUpdateSpreadsheetRequest
        {
            Requests = new List<Request> { request }
        };

        await Service.Spreadsheets.BatchUpdate(update, GSheetId).ExecuteAsync();
    }

    /// <summary>
    /// Create a new tab for statement
    /// </summary>
    /// <param name="dateTime">datetime for new name</param>
    /// <returns>title, sheetId</returns>
    public async Task<(string, int)> CreateStatementTab(DateTime dateTime)
    {
        if (Service == null)
            throw new Exception("Google sheet service has not been authenticated");

        var rand = new Random();
        var title = $"{dateTime.Year}-{dateTime.Month}";
        var id = rand.Next();

        // Add new Sheet
        var addSheetRequest = new AddSheetRequest
        {
            Properties = new SheetProperties
            {
                Title = title,
                SheetId = id
            }

        };

        var request = new Request
        {
            AddSheet = addSheetRequest
        };

        await PerformRequestAsync(request);

        return (title, id);
    }

    public async Task WriteMovements(int startRow, string tabTitle, int sheetId, ILookup<string, Movement> movements, string categoriesRange)
    {
        if (Service == null)
            throw new Exception("Google sheet service has not been authenticated");

        var sorted = movements.OrderBy(x => x.Key);
        int nRow = startRow;

        var rowsInsert = new List<ValueRange>();
        foreach (var pair in sorted)
        {
            var row = new ValueRange();
            row.Range = $"{tabTitle}!A{nRow}:D{nRow + pair.Count()}";
            row.MajorDimension = "ROWS";
            row.Values = pair.Select(x => x.GetRow(pair.Key)).ToList();
            rowsInsert.Add(row);
            nRow += pair.Count();
        }

        var update = new BatchUpdateValuesRequest
        {
            Data = rowsInsert,
            ValueInputOption = "USER_ENTERED"
        };

        await Service.Spreadsheets.Values.BatchUpdate(update, GSheetId).ExecuteAsync();

        // resize
        var autoResize = new AutoResizeDimensionsRequest
        {
            Dimensions = new DimensionRange
            {
                SheetId = sheetId,
                Dimension = "COLUMNS",
                StartIndex = 1,
                EndIndex = 3
            }
        };

        var requestSize = new Request
        {
            AutoResizeDimensions = autoResize,
        };

        await PerformRequestAsync(requestSize);

        // validation
        var dataValidation = new SetDataValidationRequest
        {
            Range = new GridRange
            {
                SheetId = sheetId,
                StartColumnIndex = 3,
                EndColumnIndex = 4,
                StartRowIndex = 0,
                EndRowIndex = nRow - 1
            },
            Rule = new DataValidationRule
            {
                Condition = new BooleanCondition
                {
                    Type = "ONE_OF_RANGE",
                    Values = new List<ConditionValue>{ new ConditionValue
                    {
                        UserEnteredValue = $"={categoriesRange}"
                    } }
                },
                Strict = true
            }
        };

        var requestValidation = new Request
        {
            SetDataValidation = dataValidation
        };

        await PerformRequestAsync(requestValidation);

    }
}
