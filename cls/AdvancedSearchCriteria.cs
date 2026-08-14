namespace Adressen.cls;

public enum SearchMode
{
    Contains, StartsWith, Exact
}

public enum SearchLogic
{
    And, Or
}

/// <summary>
/// Hält die Kriterien einer erweiterten Suche und liefert typisierte Prädikate
/// für lokale Adressen und Google-Kontakte.
/// PLZ-Bereich und Geburtsjahr-Bereich sind immer UND-Bedingungen (unabhängig von <see cref="Logic"/>).
/// </summary>
public class AdvancedSearchCriteria
{
    // ── Textfelder ───────────────────────────────────────────────────────────
    public string? Vorname
    {
        get; init;
    }
    public string? Nachname
    {
        get; init;
    }
    public string? Nickname
    {
        get; init;
    }
    public string? Praefix
    {
        get; init;
    }   // "Titel" in der UI
    public string? Anrede
    {
        get; init;
    }
    public string? Unternehmen
    {
        get; init;
    }
    public string? Strasse
    {
        get; init;
    }
    public string? Ort
    {
        get; init;
    }

    // ── PLZ-Bereich (immer UND) ──────────────────────────────────────────────
    public string? PLZvon
    {
        get; init;
    }
    public string? PLZbis
    {
        get; init;
    }   // null = PLZvon wird mit SearchMode verglichen (kein Bereich)

    // ── Geburtsjahr-Bereich (immer UND) ─────────────────────────────────────
    public string? GeburtsjahrVon
    {
        get; init;
    }  // raw text, z.B. "19" oder "1975" – wird mit SearchMode verglichen (kein Bereich)
    public int? GeburtsjahrBis
    {
        get; init;
    }  // null = kein Bereich → SearchMode verwenden; gesetzt = Bereich von..bis

    // ── Suchmodus und Verknüpfung ────────────────────────────────────────────
    public SearchMode Mode { get; init; } = SearchMode.Contains;
    public SearchLogic Logic { get; init; } = SearchLogic.And;

    // ── Hilfseigenschaft ────────────────────────────────────────────────────
    /// <summary>Gibt true zurück, wenn kein einziges Kriterium ausgefüllt ist.</summary>
    public bool IsEmpty =>
        IsNullOrEmpty(Vorname, Nachname, Nickname, Praefix, Anrede,
                      Unternehmen, Strasse, Ort, PLZvon, GeburtsjahrVon) &&
        GeburtsjahrBis is null;

    // ── Prädikate ────────────────────────────────────────────────────────────

    /// <summary>Liefert ein Prädikat für lokale <see cref="Adresse"/>-Objekte.</summary>
    public Func<Adresse, bool> BuildAdressePredicate() =>
        a => MatchesTextFields(
                 (Vorname, a.Vorname),
                 (Nachname, a.Nachname),
                 (Nickname, a.Nickname),
                 (Praefix, a.Praefix),
                 (Anrede, a.Anrede),
                 (Unternehmen, a.Unternehmen),
                 (Strasse, a.Strasse),
                 (Ort, a.Ort))
             && MatchesPLZ(a.PLZ)
             && MatchesGeburtstag(a.Geburtstag);

    /// <summary>Liefert ein Prädikat für Google-<see cref="Contact"/>-Objekte.</summary>
    public Func<Contact, bool> BuildContactPredicate() =>
        c => MatchesTextFields(
                 (Vorname, c.Vorname),
                 (Nachname, c.Nachname),
                 (Nickname, c.Nickname),
                 (Praefix, c.Praefix),
                 (Anrede, c.Anrede),
                 (Unternehmen, c.Unternehmen),
                 (Strasse, c.Strasse),
                 (Ort, c.Ort))
             && MatchesPLZ(c.PLZ)
             && MatchesGeburtstag(c.Geburtstag);

    // ── Private Matching-Logik ───────────────────────────────────────────────

    /// <summary>
    /// Wertet nur Paare aus, bei denen das Kriterium nicht leer ist.
    /// Leere Kriterien werden ignoriert (nicht als "Feld muss leer sein" interpretiert).
    /// </summary>
    private bool MatchesTextFields(params (string? Criterion, string? Value)[] pairs)
    {
        var active = pairs.Where(p => !string.IsNullOrWhiteSpace(p.Criterion)).ToArray();
        if (active.Length == 0) { return true; }  // Keine Textkriterien → kein Filter
        return Logic == SearchLogic.And ? active.All(p => MatchText(p.Criterion!, p.Value)) : active.Any(p => MatchText(p.Criterion!, p.Value));
    }

    private bool MatchText(string criterion, string? value)
    {
        if (string.IsNullOrEmpty(value)) { return false; }
        return Mode switch
        {
            SearchMode.Contains => value.Contains(criterion, StringComparison.OrdinalIgnoreCase),
            SearchMode.StartsWith => value.StartsWith(criterion, StringComparison.OrdinalIgnoreCase),
            SearchMode.Exact => string.Equals(value.Trim(), criterion.Trim(), StringComparison.OrdinalIgnoreCase),
            _ => false
        };
    }

    /// <summary>
    /// PLZ-Bereich: immer UND (unabhängig von <see cref="Logic"/>).
    /// Ist PLZbis gesetzt, wird ein lexikographischer Bereichsvergleich durchgeführt
    /// (funktioniert korrekt für gleichlange PLZ – DE 5-stellig, AT/CH 4-stellig).
    /// Ist PLZbis nicht gesetzt, wird PLZvon mit dem konfigurierten <see cref="Mode"/> verglichen.
    /// </summary>
    private bool MatchesPLZ(string? plz)
    {
        if (string.IsNullOrWhiteSpace(PLZvon)) { return true; }
        if (string.IsNullOrWhiteSpace(plz)) { return false; }

        if (!string.IsNullOrWhiteSpace(PLZbis))
        {
            // Bereichsmodus: von ≤ PLZ ≤ bis (lexikographisch, setzt gleichlange PLZ voraus)
            var v = plz.Trim();
            return string.CompareOrdinal(v, PLZvon.Trim()) >= 0
                && string.CompareOrdinal(v, PLZbis.Trim()) <= 0;
        }

        // Kein Bereich: SearchMode (Contains / StartsWith / Exact) anwenden
        return MatchText(PLZvon, plz);
    }

    /// <summary>
    /// Geburtsjahr-Bereich: immer UND (unabhängig von <see cref="Logic"/>).
    /// Ist GeburtsjahrBis gesetzt, wird ein ganzzahliger Bereichsvergleich durchgeführt.
    /// Ist GeburtsjahrBis nicht gesetzt, wird GeburtsjahrVon mit dem konfigurierten <see cref="Mode"/> verglichen.
    /// </summary>
    private bool MatchesGeburtstag(DateOnly? geburtstag)
    {
        if (string.IsNullOrWhiteSpace(GeburtsjahrVon)) { return true; }   // kein Jahresfilter
        if (!geburtstag.HasValue) { return false; }

        if (GeburtsjahrBis.HasValue && int.TryParse(GeburtsjahrVon, out var von))
        {
            // Bereichsmodus: von ≤ Jahr ≤ bis (setzt vollständige 4-stellige Jahresangabe voraus)
            return geburtstag.Value.Year >= von && geburtstag.Value.Year <= GeburtsjahrBis.Value;
        }

        // Kein Bereich: SearchMode (Contains / StartsWith / Exact) auf Jahreszahl als String anwenden
        return MatchText(GeburtsjahrVon, geburtstag.Value.Year.ToString());
    }

    private static bool IsNullOrEmpty(params string?[] values) => values.All(string.IsNullOrWhiteSpace);
}
