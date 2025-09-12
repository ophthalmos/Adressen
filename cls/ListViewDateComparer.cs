using System.Collections;
using System.Globalization;

namespace Adressen.cls;

public class ListViewItemComparer : IComparer
{
    private readonly int _column;
    private readonly SortOrder _order;

    public ListViewItemComparer()
    {
        _column = 2;
        _order = SortOrder.Descending;
    }

    public ListViewItemComparer(int column, SortOrder order)
    {
        _column = column;
        _order = order;
    }

    public int Compare(object? x, object? y)
    {
        var result = -1;
        var itemX = x as ListViewItem;
        var itemY = y as ListViewItem;
        var textX = itemX?.SubItems[_column].Text;
        var textY = itemY?.SubItems[_column].Text;
        if (_column == 2)
        {
            var successX = DateTime.TryParse(textX, new CultureInfo("de-DE"), out var dateX);
            var successY = DateTime.TryParse(textY, new CultureInfo("de-DE"), out var dateY);
            if (successX && successY) { result = DateTime.Compare(dateX, dateY); }
        }
        else if (_column == 1)
        {
            var sizeX = ParseSize(textX ?? string.Empty);
            var sizeY = ParseSize(textY ?? string.Empty);
            result = sizeX.CompareTo(sizeY);
        }
        else { result = string.Compare(textX, textY); }
        return _order == SortOrder.Descending ? -result : result;

    }

    private static long ParseSize(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) { return 0; }
        var parts = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2 || !double.TryParse(parts[0], NumberStyles.Float, CultureInfo.CurrentCulture, out var value)) { return 0; }
        var unit = parts[1].ToUpperInvariant();
        var multiplier = unit switch
        {
            "B" => 1L,
            "KB" => 1024L,
            "MB" => 1024L * 1024L,
            "GB" => 1024L * 1024L * 1024L,
            "TB" => 1024L * 1024L * 1024L * 1024L,
            _ => 1L
        };
        return (long)(value * multiplier);
    }
}