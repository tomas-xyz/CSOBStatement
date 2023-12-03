namespace tomxyz.csob;

public class Movement
{
    public DateTime Date { get; set; }

    public double Amount { get; set; }

    public string Account { get; set; } = string.Empty;

    public long SpecificSymbol { get; set; }

    public long VariableSymbol { get; set; }
    public string AccountId { get; set; } = string.Empty;

    public IEnumerable<string> Messages { get; set; } = new List<string>();

    public string GetCategory(IEnumerable<Rule> rules)
    {
        foreach (var rule in rules)
        {
            if (rule.MovementFit(this))
                return rule.Category;
        }

        return string.Empty;
    }

    public IList<object> GetRow(string category)
    {
        return new List<object>()
        {
            Date.ToString("dd.MM yyyy"),
            RowId(),
            Amount.ToString(),
            category,
            "FALSE"
        };
    }

    protected string RowId()
    {
        var id = string.Empty;

        if (!string.IsNullOrEmpty(AccountId))
            id += AccountId + ", ";

        if (Messages.Any() && !string.IsNullOrEmpty(Messages.First()))
            id += Messages.First();

        if (string.IsNullOrEmpty(id))
            id = Messages.Where(x => !string.IsNullOrEmpty(x)).First() ?? "Žádný popis";

        return id;
    }
}
