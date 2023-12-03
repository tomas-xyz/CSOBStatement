
namespace tomxyz.csob;
public class CategoriesComparer : IComparer<string>
{
    private Dictionary<string, int> Indexes { get; }
    public CategoriesComparer(IEnumerable<string> categories)
    {
        Indexes = categories.ToDictionary(x => x, x => categories.ToList().IndexOf(x));
    }

    public int Compare(string? x, string? y)
    {
        var first = Indexes[x ?? throw new Exception("Category is null during comparing")];
        var second = Indexes[y ?? throw new Exception("Category is null during comparing")];

        return first.CompareTo(second);
    }
}
