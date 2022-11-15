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

    protected static IList<object> GetRow(IEnumerable<string> values)
    {
        return values.Select(x => (object) x).ToList();        
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
        var title = $"{dateTime.Year}-{dateTime.Month.ToString("00")}";
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

    /// <summary>
    /// Format cell range
    /// </summary>
    /// <param name="sheetId">sheet id</param>
    /// <param name="startColumn">start column index</param>
    /// <param name="endColumn">end column index</param>
    /// <param name="startRow">start row index</param>
    /// <param name="endRow">end row index</param>
    /// <param name="bold">make it bold</param>
    /// <param name="center">center it</param>
    /// <param name="currency">format number as currency</param>
    /// <returns></returns>
    protected async Task FormatCellsAsync(int sheetId, int startColumn, int endColumn, int startRow, int endRow, bool bold, bool center, bool currency)
    {
        var request = new RepeatCellRequest
        {
            Range = new GridRange
            {
                SheetId = sheetId,
                StartColumnIndex = startColumn,
                EndColumnIndex = endColumn,
                StartRowIndex = startRow,
                EndRowIndex = endRow
            },
            Cell = new CellData
            {
                UserEnteredFormat = new CellFormat
                {
                    TextFormat = new TextFormat
                    {
                        Bold = bold                        
                    },
                    HorizontalAlignment = (center ? "CENTER" : "LEFT"),
                    NumberFormat = new NumberFormat
                    {
                        Type = currency ? "CURRENCY" : "NUMBER_FORMAT_TYPE_UNSPECIFIED"
                    }
                }
            },
            Fields = "UserEnteredFormat"
        };


        var requestFormat = new Request
        {
            RepeatCell = request
        };

        await PerformRequestAsync(requestFormat);
    }

    public async Task<int> WriteSummaryAsync(string tabTitle, int sheetId, Statement statement)
    {
        if (Service == null)
            throw new Exception("Google sheet service has not been authenticated");

        var values = new List<IList<object>>
        {
            GetRow(new []{"Jméno:", statement.Name }),
            GetRow(new []{"Účet:", statement.Account}),
            GetRow(new []{"Od:", statement.DateFrom.ToString("dd.MM yyyy")}),
            GetRow(new []{"Do:", statement.DateTo.ToString("dd.MM yyyy")}),
            GetRow(new []{"Počáteční stav", statement.StartAmount.ToString()}),
            GetRow(new []{"Koncový stav", (statement.StartAmount + statement.Plus + statement.Minus).ToString()}),            
            GetRow(new []{"Příjmy", (statement.Plus).ToString()}),
            GetRow(new []{"Výdaje", (statement.Minus).ToString()}),
            GetRow(new []{"Bilance", (statement.Plus + statement.Minus).ToString()}),
        };

        var rowsInsert = new List<ValueRange>();
        rowsInsert.Add(
            new ValueRange
            {
                Range = $"{tabTitle}!A1:B10",
                MajorDimension = "ROWS",
                Values = values
            });

        var update = new BatchUpdateValuesRequest
        {
            Data = rowsInsert,
            ValueInputOption = "USER_ENTERED"
        };

        await Service.Spreadsheets.Values.BatchUpdate(update, GSheetId).ExecuteAsync();

        // left column - bold
        await FormatCellsAsync(
            sheetId,
            0,
            1,
            0,
            values.Count,
            true,
            false,
            false);

        // right column currency and center
        await FormatCellsAsync(
            sheetId,
            1,
            2,
            0,
            values.Count,
            false,
            true,
            true);

        // bilance bold
        await FormatCellsAsync(
            sheetId,
            1,
            2,
            values.Count - 1,
            values.Count,
            true,
            true,
            true);

        return values.Count() + 1;
    }

    public async Task WriteMovements(int startRow, string tabTitle, int sheetId, ILookup<string, Movement> movements, string categoriesRange)
    {
        if (Service == null)
            throw new Exception("Google sheet service has not been authenticated");

        var sorted = movements.OrderBy(x => x.Key);
        int nRow = startRow;

        var rowsInsert = new List<ValueRange>();

        rowsInsert.Add(
           new ValueRange
           {
               Range = $"{tabTitle}!A{nRow}:D{nRow}",
               MajorDimension = "ROWS",
               Values =new List<IList<object>> { GetRow(new[] { "Datum", "Místo / zpráva", "Částka", "Kategorie" }) }
           });

        var startValidatedRow = nRow++;
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

        // header
        await FormatCellsAsync(
            sheetId,
            0,
            4,
            startValidatedRow - 1,
            startValidatedRow,
            true,
            true,
            false);

        // center date
        await FormatCellsAsync(
            sheetId,
            0,
            1,
            startValidatedRow,
             nRow - 1,
            false,
            true,
            false);

        // center amount  and category and format amount as currency
        await FormatCellsAsync(
            sheetId,
            2,
            4,
            startValidatedRow,
             nRow - 1,
            false,
            true,
            true);

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
                StartRowIndex = startValidatedRow,
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
