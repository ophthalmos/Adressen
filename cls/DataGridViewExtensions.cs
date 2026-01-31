namespace Adressen.cls;

internal static class DataGridViewExtensions
{
    public static (bool[] HideArr, int[] WidthArr) GetUpdatedSettings(this DataGridView dgv, bool[] savedHide, int[] savedWidths, bool[] defaultHide, int[] defaultWidths)
    {
        var count = dgv.Columns.Count;
        if (count == 0) { return (savedHide, savedWidths); }
        bool[] finalHide = (savedHide.Length == count) ? [.. savedHide] : [.. defaultHide];
        int[] finalWidths = (savedWidths.Length == count) ? [.. savedWidths] : [.. defaultWidths];
        for (var i = 0; i < count; i++)
        {
            var colName = dgv.Columns[i].Name;
            var w = i < finalWidths.Length ? finalWidths[i] : (colName == "Nachname" ? 200 : 100); // Breite setzen
            dgv.Columns[i].Width = Math.Max(20, w); // Mindestbreite 20px
            finalWidths[i] = dgv.Columns[i].Width;
            dgv.Columns[i].Visible = !finalHide[i];
        }
        return (finalHide, finalWidths);
    }
}