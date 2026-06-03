namespace Adressen.cls;

internal class TextBoxSearchManager
{
    public static List<string> SearchHistory { get; } = [];

    private string? _searchString = string.Empty;
    private int _searchStart = -1;

    public static bool CaseChecked { get; set; } = false;  // Statisch, damit die Einstellung über alle Formular-Instanzen hinweg erhalten bleibt

    public bool HasSearchTerm => !string.IsNullOrEmpty(_searchString);

    public void ShowSearchDialogAndSearch(TextBox tb)
    {
        var selectedText = tb.SelectedText;
        if (!string.IsNullOrEmpty(selectedText)) { _searchString = selectedText; }
        using var f = new FrmTextBoxSearch(_searchString ?? string.Empty, CaseChecked);
        if (f.ShowDialog() == DialogResult.OK)
        {
            _searchString = f.SearchText;
            CaseChecked = f.MatchCase;
            _searchStart = -1; // Suche von vorne (oder ab Cursor) starten
            PerformSearch(tb);
        }
    }

    public void FindNext(TextBox tb) => PerformSearch(tb);

    private void PerformSearch(TextBox tb)
    {
        if (string.IsNullOrEmpty(_searchString)) { return; }
        UpdateSearchHistory(_searchString);
        var comparison = CaseChecked ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var startIndex = _searchStart + 1;
        if (startIndex >= tb.TextLength) { startIndex = 0; }
        var matchIndex = tb.Text.IndexOf(_searchString, startIndex, comparison);
       if (matchIndex == -1 && startIndex > 0)        {            matchIndex = tb.Text.IndexOf(_searchString, 0, comparison); }  // Wrap-Around: Wenn nichts gefunden wurde, aber wir nicht von Position 0 gestartet sind
        if (matchIndex != -1)
        {
            _searchStart = matchIndex;
            tb.SelectionStart = _searchStart;
            tb.SelectionLength = _searchString.Length;
            tb.Select();
            tb.ScrollToCaret();
        }
        else
        {
            Utils.MsgTaskDlg(tb.Handle, "Suche in Notizen", "Der Suchtext wurde nicht gefunden.", TaskDialogIcon.Information);
            _searchStart = -1;
        }
    }

    private static void UpdateSearchHistory(string searchTerm)
    {
        if (SearchHistory != null)
        {
            SearchHistory.Remove(searchTerm);  // Vorhandenen Eintrag entfernen, um ihn nach oben zu setzen
            SearchHistory.Insert(0, searchTerm);
            while (SearchHistory.Count > 10) { SearchHistory.RemoveAt(SearchHistory.Count - 1); }  // Limit auf 10 Einträge erzwingen
        }
    }
}