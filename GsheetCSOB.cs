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

    protected static IList<object> GetRow(IEnumerable<object> values)
    {
        return values.Select(x => x).ToList();        
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
    /// <param name="cellFormat">format number</param>
    /// <param name="checkBox">format as checkbox</param>
    /// <returns></returns>
    protected async Task FormatCellsAsync(int sheetId, int startColumn, int endColumn, int startRow, int endRow, bool bold, bool center, CellFormat cellFormat = CellFormat.NUMBER_FORMAT_TYPE_UNSPECIFIED)
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
                UserEnteredFormat = new Google.Apis.Sheets.v4.Data.CellFormat
                {
                    TextFormat = new TextFormat
                    {
                        Bold = bold                        
                    },
                    HorizontalAlignment = (center ? "CENTER" : "LEFT"),
                    NumberFormat = new NumberFormat
                    {
                        Type = cellFormat.ToString(),
                        Pattern = cellFormat == CellFormat.PERCENT ? "#0.00%" : ""
                    }
                },
            },
            Fields = "UserEnteredFormat"
        };

        var requestFormat = new Request
        {
            RepeatCell = request
        };

        await PerformRequestAsync(requestFormat);
    }

    protected async Task CheckBoxCellsAsync(int sheetId, int column, int startRow, int endRow)
    {
        var request = new RepeatCellRequest
        {
            Range = new GridRange
            {
                SheetId = sheetId,
                StartColumnIndex = column,
                EndColumnIndex = column+1,
                StartRowIndex = startRow,
                EndRowIndex = endRow
            },
            Cell = new CellData
            {
                DataValidation = new DataValidationRule
                {
                    Condition = new BooleanCondition
                    {
                        Type = "BOOLEAN"
                    }
                }
            },
            Fields = "DataValidation"
        };

        var requestFormat = new Request
        {
            RepeatCell = request
        };

        await PerformRequestAsync(requestFormat);
    }

    /// <summary>
    /// Write summary information about the statement
    /// </summary>
    /// <param name="tabTitle">tab title</param>
    /// <param name="sheetId">sheet id</param>
    /// <param name="statement">statement</param>
    /// <returns>first empty row</returns>
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
            GetRow(new []{"Příjmy", "=SUMIFS(C12:C;C12:C;\">0\";E12:E;NEPRAVDA)"}),
            GetRow(new []{"Výdaje","=SUMIFS(C12:C;C12:C;\"<0\";E12:E;NEPRAVDA)" }),
            GetRow(new []{"Bilance", "=B7+B8" }),
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
            CellFormat.CURRENCY);

        // bilance bold
        await FormatCellsAsync(
            sheetId,
            1,
            2,
            values.Count - 1,
            values.Count,
            true,
            true,
            CellFormat.CURRENCY);

        return values.Count() + 1;
    }

    /// <summary>
    /// Write movements with data validation
    /// </summary>
    /// <param name="startRow">first row for movements</param>
    /// <param name="tabTitle">tab title</param>
    /// <param name="sheetId">sheet id</param>
    /// <param name="movements">movements</param>
    /// <param name="categoriesRange">categories</param>
    public async Task WriteMovements(int startRow, string tabTitle, int sheetId, IOrderedEnumerable<IGrouping<string, Movement>> movements, string categoriesRange)
    {
        if (Service == null)
            throw new Exception("Google sheet service has not been authenticated");

        int nRow = startRow;

        var rowsInsert = new List<ValueRange>();

        rowsInsert.Add(
           new ValueRange
           {
               Range = $"{tabTitle}!A{nRow}:E{nRow}",
               MajorDimension = "ROWS",
               Values =new List<IList<object>> { GetRow(new[] { "Datum", "Místo / zpráva", "Částka", "Kategorie", "Ignorovat" }) }
           });

        var startValidatedRow = nRow++;
        foreach (var pair in movements)
        {
            var row = new ValueRange();
            row.Range = $"{tabTitle}!A{nRow}:E{nRow + pair.Count()}";
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
            5,
            startValidatedRow - 1,
            startValidatedRow,
            true,
            true);

        // center date
        await FormatCellsAsync(
            sheetId,
            0,
            1,
            startValidatedRow,
             nRow - 1,
            false,
            true);

        // center amount  and category and format amount as currency
        await FormatCellsAsync(
            sheetId,
            2,
            4,
            startValidatedRow,
             nRow - 1,
            false,
            true,
            CellFormat.CURRENCY);

        // "Ignore" column
        await CheckBoxCellsAsync(
            sheetId,
            4,
            startValidatedRow,
             nRow - 1);

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

        // resize
        var autoResize = new AutoResizeDimensionsRequest
        {
            Dimensions = new DimensionRange
            {
                SheetId = sheetId,
                Dimension = "COLUMNS",
                StartIndex = 0,
                EndIndex = 4
            }
        };

        var requestSize = new Request
        {
            AutoResizeDimensions = autoResize,
        };

        await PerformRequestAsync(requestSize);
    }

    public async Task WriteStatisticsAsync(string tabTitle, int sheetId, IOrderedEnumerable<IGrouping<string, Movement>> movements, string range, IEnumerable<string> categories)
    {
        if (Service == null)
            throw new Exception("Google sheet service has not been authenticated");

        var header = new List<IList<object>>
        {
            GetRow(new []{"Kategorie", "Suma", "Procento výdajů / příjmů"})
        };

        var rowsInsert = new List<ValueRange>();
        rowsInsert.Add(
            new ValueRange
            {
                Range = $"{tabTitle}!G2:I2",
                MajorDimension = "ROWS",
                Values = header
            });

        var update = new BatchUpdateValuesRequest
        {
            Data = rowsInsert,
            ValueInputOption = "USER_ENTERED"
        };

        await Service.Spreadsheets.Values.BatchUpdate(update, GSheetId).ExecuteAsync();


        // write links to categories and sumifs formulas
        var rowsStats = new List<ValueRange>();
        var baseRange = range.Substring(0, range.IndexOf(':'));
        for (var i = 1; i < categories.Count() + 1; i++)
        {
            var values = new List<IList<object>>
            {                
                GetRow(new []{$"={baseRange}{i}", $"=SUMIFS(C12:C;D12:D;G{2+i};E12:E;FALSE)"})
            };

            rowsStats.Add(
                new ValueRange
                {
                    Range = $"{tabTitle}!G{2+i}:I{2+i}",
                    MajorDimension = "ROWS",
                    Values = values
                });
        }

        var valuesInsert = new BatchUpdateValuesRequest
        {
            Data = rowsStats,
            ValueInputOption = "USER_ENTERED"
        };

        await Service.Spreadsheets.Values.BatchUpdate(valuesInsert, GSheetId).ExecuteAsync();

        // get sums to determine positive/negative numbers
        var request = Service.Spreadsheets.Values.Get(GSheetId, $"{tabTitle}!H3:H{3 + categories.Count()}");
        var sums = (await request.ExecuteAsync(default)).Values.SelectMany(x => x).Select(i => double.Parse((string)i));
        var rowsPercs = new List<ValueRange>();

        int n = 1;
        foreach (var sum in sums)
        {
            var values = new List<IList<object>>
            {                
                sum < 0 ? GetRow(new []{$"=H{2 + n}/$B$8"}) : GetRow(new []{$"=H{2 + n}/$B$7"})
            };

            rowsPercs.Add(
                new ValueRange
                {
                    Range = $"{tabTitle}!I{2 + n}",
                    MajorDimension = "ROWS",
                    Values = values
                });
            
            ++n;
        }

        var valuesPercs = new BatchUpdateValuesRequest
        {
            Data = rowsPercs,
            ValueInputOption = "USER_ENTERED"
        };

        await Service.Spreadsheets.Values.BatchUpdate(valuesPercs, GSheetId).ExecuteAsync();

        // header column - bold and center
        await FormatCellsAsync(
            sheetId,
            5,
            9,
            1,
            2,
            true,
            true);

        // second header column - percentage
        await FormatCellsAsync(
            sheetId,
            7,
            8,
            2,
            100,
            false,
            true,
            CellFormat.CURRENCY);

        // second header column - percentage
        await FormatCellsAsync(
            sheetId,
            8,
            9,
            2,
            100,
            false,
            true,
            CellFormat.PERCENT);

        // resize
        var autoResize = new AutoResizeDimensionsRequest
        {
            Dimensions = new DimensionRange
            {
                SheetId = sheetId,
                Dimension = "COLUMNS",
                StartIndex = 6,
                EndIndex = 9
            }
        };

        var requestSize = new Request
        {
            AutoResizeDimensions = autoResize,
        };

        await PerformRequestAsync(requestSize);
    }

}
