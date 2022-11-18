using System.Text.RegularExpressions;

namespace tomxyz.csob;
public class Rule
{
    private string? MessageSubstring { get; } = null;
    private Regex? MessageRegex { get; } = null;
    private string? Account { get; } = null;
    private int? Amount { get; } = null;
    internal string Category { get; }

    public Rule(string account, string message, string amount, string category)
    {
        if (!string.IsNullOrEmpty(message))
        {
            MessageSubstring = message.ToLower();
            try
            {
                MessageRegex = new Regex(message, RegexOptions.IgnoreCase);
            }
            catch (Exception)
            {
                MessageRegex = null;
            }
        }

        if (!string.IsNullOrEmpty(account))
            Account = account;

        if (!string.IsNullOrEmpty(amount))
            Amount = int.Parse(amount);

        Category = category;    
    }

    public bool MovementFit(Movement movement)
    {
        if (Account != null && Account != movement.Account)
            return false;

        if (MessageSubstring != null)
        {
            if (!movement.Messages.Any(x => x.ToLower().Contains(MessageSubstring)))
            {
                if (MessageRegex == null || !movement.Messages.Any(x => MessageRegex.IsMatch(x)))
                    return false;
            }
        }

        if (Amount != null)
        {
            if (Amount < 0 && movement.Amount >= Amount)
                return false;
            else if (Amount >= 0 && movement.Amount < Amount)
                return false;
        }

        return true;
    }
}
