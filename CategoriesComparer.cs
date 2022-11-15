
namespace tomxyz.csob;
public class CategoriesComparer  : IComparer<string>
{
    private Dictionary<string, int> Indexes { get; }
    public CategoriesComparer(IEnumerable<string> categories)
    {
        Indexes = categories.ToDictionary(x => x, x => categories.ToList().IndexOf(x));
    }

    public int Compare(string? x, string? y)
    {
        var first = Indexes[x];
        var second = Indexes[y];

        return first.CompareTo(second);         
    }
}
