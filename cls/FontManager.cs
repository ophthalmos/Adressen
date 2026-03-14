namespace Adressen.cls;

public static class FontManager
{
    private static Task<List<string>>? _preloadTask;
    private static readonly Dictionary<string, Font> _fontCache = []; // Der globale Cache

    public static void StartPreloading() => _preloadTask ??= Task.Run(LoadFonts);

    public static List<string> GetValidFonts()
    {
        if (_preloadTask == null) { StartPreloading(); }
        return _preloadTask!.Result;
    }

    public static Font GetDisplayFont(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName)) { fontName = "Segoe UI"; }

        if (!_fontCache.TryGetValue(fontName, out var displayFont))
        {
            try { displayFont = new Font(fontName, 10f, FontStyle.Regular, GraphicsUnit.Point); }
            catch { displayFont = new Font("Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point); }
            _fontCache[fontName] = displayFont;
        }
        return displayFont;
    }

    public static void Cleanup()   // Gibt alle GDI-Ressourcen am Ende des Programms frei
    {
        foreach (var font in _fontCache.Values) { font.Dispose(); }
        _fontCache.Clear();
    }

    private static List<string> LoadFonts()
    {
        var validFontNames = new List<string>();
        var blockList = new[] { "icon", "mdl2", "emoji", "symbol", "dingbat", "marlett", "webdings", "wingdings", "math", "hollow" };
        foreach (var family in FontFamily.Families)
        {
            var nameLower = family.Name.ToLower();
            if (blockList.Any(nameLower.Contains)) { continue; }
            try
            {
                using var testFont = new Font(family, 10f, FontStyle.Regular);
                if (testFont.GdiCharSet == 0 || testFont.GdiCharSet == 1) { validFontNames.Add(family.Name); }
            }
            catch { }   // Ignorieren
        }
        return validFontNames;
    }
}