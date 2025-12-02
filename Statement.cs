using System.Text;
using System.Xml;
using System.Xml.Serialization;
using tomxyz.csob.xml;

namespace tomxyz.csob;

public class Statement
{
    public string Name { get; }

    public string Account { get; }

    public DateTime DateFrom { get; }

    public DateTime DateTo { get; }

    public double StartAmount { get; set; }

    public double Plus { get; set; }

    public double Minus { get; set; }

    public IEnumerable<Movement> Movements { get; set; }

    public Dictionary<string, IEnumerable<Movement>> CategorizedMovements;

    public Statement(string filePath)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var reader = XmlReader.Create(filePath);
        if (!reader.ReadToFollowing("FINSTA03"))
            throw new Exception("Element 'FINSTA03' has not been found in input xml");

        XmlSerializer serializer = new XmlSerializer(typeof(StatementXml));
        var statement = (StatementXml)(serializer.Deserialize(reader) ?? throw new Exception("Unable to deserialize statement from xml"));
        if (statement == null)
            throw new Exception("Deserialization of statement has failed");

        CategorizedMovements = new Dictionary<string, IEnumerable<Movement>>();
        Name = statement.Name;
        Account = statement.Account;
        DateFrom = DateTime.Parse(statement.DateFrom);
        DateTo = DateTime.Parse(statement.DateTo);
        StartAmount = double.Parse(statement.StartAmount);
        Plus = double.Parse(statement.Plus.Substring(statement.Plus.IndexOf('=') + 1));
        Minus = -double.Parse(statement.Minus.Substring(statement.Minus.IndexOf('=') + 1));
        Movements = statement.MovementsXml.Select(x => new Movement
        {
            AccountType = statement.AccountType,
            Date = DateTime.Parse(x.DateString),
            Amount = double.Parse(x.Amount),
            Account = x.Account,
            BankId = x.BankId,
            AccountId = x.AccountId,
            SpecificSymbol = long.Parse(x.SpecificSymbol),
            VariableSymbol = long.Parse(x.VariableSymbol),
            Messages = new List<string>
            {
                x.Message1,
                x.Message2,
                x.Message3,
                x.Message4,
                x.Message5,
                x.Message6,
                x.Message7,
                x.Message8
            }.Where(m => !string.IsNullOrEmpty(m))
        });
    }
}


