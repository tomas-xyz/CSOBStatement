namespace tomxyz.csob;

public class Movement
{
    public DateTime Date { get; set; }

    public double Amount { get; set; }

    public string Account { get; set; }

    public long SpecificSymbol { get; set; }

    public long VariableSymbol { get; set; }

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
            Messages.Where(x => !string.IsNullOrEmpty(x)).FirstOrDefault(),
            Amount.ToString(),
            category
        };
    }
}
