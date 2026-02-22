using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;

namespace Adressen.cls;

public static class DbChangeAnalyzer
{
    //public record ChangeAnalysisResult(bool HasChanges, List<EntityEntry> RealChanges, string DialogHeading, string DialogText);
    public record ChangeAnalysisResult(bool HasChanges, List<EntityEntry> RealChanges, string DialogHeading, string DialogText, string ExpanderText);

    //public static ChangeAnalysisResult AnalyzeChanges(DbContext? context)
    //{
    //    if (context == null) { return new ChangeAnalysisResult(false, [], string.Empty, string.Empty); }

    //    context.ChangeTracker.DetectChanges();

    //    // Wir holen alle echten Änderungen (ignoriert reine Whitespace-Änderungen)
    //    var allRealChanges = context.ChangeTracker.Entries().Where(IsEntryReallyChanged).ToList();

    //    if (allRealChanges.Count == 0) { return new ChangeAnalysisResult(false, [], string.Empty, string.Empty); }

    //    // --- 1. Echte Adress-Änderungen (Namen, Telefon etc.) ---
    //    var changedAddresses = allRealChanges
    //        .Where(e => e.Metadata.ClrType == typeof(Adresse))
    //        .Select(e => (Adresse)e.Entity)
    //        .ToHashSet(); // HashSet verhindert Duplikate

    //    // --- 2. Indirekte Adress-Änderungen (Gruppen-Zuordnungen) ---
    //    // Wir suchen Einträge, die keine Klasse haben (Schatten-Entitäten) und "Gruppe" im Namen tragen
    //    var shadowEntries = allRealChanges.Where(e => e.Metadata.ClrType == null && e.Metadata.Name.Contains("Gruppe"));

    //    foreach (var shadow in shadowEntries)
    //    {
    //        // Wir suchen den Fremdschlüssel, der zur Adresse zeigt
    //        foreach (var fk in shadow.Metadata.GetForeignKeys())
    //        {
    //            if (fk.PrincipalEntityType.ClrType == typeof(Adresse))
    //            {
    //                // Wert des Fremdschlüssels (AdressId) aus der Schatten-Entität lesen
    //                // Da es meist nur eine Property für den FK gibt (AdressId), nehmen wir die erste.
    //                // Wir prüfen die Anzahl und greifen direkt auf den Index 0 zu
    //                var fkProp = fk.Properties.Count > 0 ? fk.Properties[0] : null;
    //                if (fkProp != null && shadow.CurrentValues[fkProp] is int addressId)
    //                {
    //                    var addr = context.Set<Adresse>().Local.FirstOrDefault(a => a.Id == addressId);
    //                    if (addr != null) { changedAddresses.Add(addr); }
    //                }
    //            }
    //        }
    //    }

    //    // --- Text-Generierung ---
    //    var changedAddressNames = changedAddresses
    //        .Select(a =>
    //        {
    //            var fullName = $"{a.Vorname} {a.Nachname}".Trim();
    //            if (!string.IsNullOrWhiteSpace(fullName)) { return $"• {fullName}"; }
    //            if (!string.IsNullOrWhiteSpace(a.Unternehmen)) { return $"• {a.Unternehmen}"; }
    //            return "• [N. n.]";
    //        })
    //        .OrderBy(n => n)
    //        .ToList();
    //    var addressChangesCount = changedAddressNames.Count;

    //    // 2. Andere Entitäten filtern
    //    var otherChanges = allRealChanges.Where(e => e.Metadata.ClrType != typeof(Adresse)).ToList();
    //    var otherChangesCount = otherChanges.Count;

    //    var groupCount = otherChanges.Count(e =>
    //        e.Metadata.ClrType == typeof(Gruppe) ||
    //        (e.Metadata.ClrType == null && e.Metadata.Name.Contains("Gruppe"))); // Erkennt "AdresseGruppe" Join-Table ohne Klasse

    //    var photoCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Foto));
    //    var docCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Dokument));

    //    var heading = addressChangesCount > 0
    //    ? (addressChangesCount == 1 ? "Möchten Sie die Änderung speichern?" : "Möchten Sie die Änderungen speichern?")
    //    : "Änderungen speichern?";

    //    var text = string.Empty;
    //    if (addressChangesCount > 0)
    //    {
    //        heading = addressChangesCount == 1 ? "Möchten Sie die Änderung speichern?" : "Möchten Sie die Änderungen speichern?";
    //        text = addressChangesCount > 10
    //        ? string.Join(Environment.NewLine, changedAddressNames.Take(10)) + Environment.NewLine + "…"
    //        : string.Join(Environment.NewLine, changedAddressNames);
    //        // Zusatzhinweis
    //        //if (otherChangesCount > 0) { text += $"{Environment.NewLine}und {otherChangesCount} Änderungen an Zusatzdaten"; }
    //    }
    //    else
    //    {
    //        // Wenn nur Gruppen/Fotos geändert wurden, aber keine Adress-Texte
    //        heading = otherChangesCount == 1 ? "Änderung an Zusatzdaten speichern?" : $"Änderungen an {otherChangesCount} Zusatzdaten speichern?";

    //        var detailsList = new List<string>();
    //        if (groupCount > 0) { detailsList.Add(groupCount == 1 ? "einer Gruppenzuordnung" : $"{groupCount} Gruppenzuordnungen"); }
    //        if (photoCount > 0) { detailsList.Add(photoCount == 1 ? "einem Foto" : $"{photoCount} Fotos"); }
    //        if (docCount > 0) { detailsList.Add(docCount == 1 ? "einem Dokument" : $"{docCount} Dokumenten"); }
    //        //var remainder = otherChangesCount - (groupCount + photoCount + docCount);  
    //        //if (remainder > 0) { detailsList.Add(remainder == 1 ? "einem sonstigen Element" : $"{remainder} sonstigen Elementen"); }

    //        if (detailsList.Count > 0) { text = "Es wurden Änderungen an " + string.Join(", ", detailsList) + " vorgenommen."; }
    //        else { text = "Es wurden Änderungen an Zusatzdaten vorgenommen."; }
    //    }
    //    return new ChangeAnalysisResult(true, allRealChanges, heading, text);
    //}

    public static ChangeAnalysisResult AnalyzeChanges(DbContext? context)
    {
        if (context == null) { return new ChangeAnalysisResult(false, [], string.Empty, string.Empty, string.Empty); }

        context.ChangeTracker.DetectChanges();
        var allRealChanges = context.ChangeTracker.Entries().Where(IsEntryReallyChanged).ToList();

        if (allRealChanges.Count == 0) { return new ChangeAnalysisResult(false, [], string.Empty, string.Empty, string.Empty); }

        var changedAddresses = allRealChanges
            .Where(e => e.Metadata.ClrType == typeof(Adresse))
            .Select(e => (Adresse)e.Entity)
            .ToHashSet();

        // KORREKTUR: "AdresseGruppen" direkt über Name abfragen, da ClrType ein Dictionary ist
        var shadowEntries = allRealChanges.Where(e => e.Metadata.Name == "AdresseGruppen");

        foreach (var shadow in shadowEntries)
        {
            foreach (var fk in shadow.Metadata.GetForeignKeys())
            {
                if (fk.PrincipalEntityType.ClrType == typeof(Adresse))
                {
                    var fkProp = fk.Properties.Count > 0 ? fk.Properties[0] : null;
                    if (fkProp != null && shadow.CurrentValues[fkProp] is int addressId)
                    {
                        var addr = context.Set<Adresse>().Local.FirstOrDefault(a => a.Id == addressId);
                        if (addr != null) { changedAddresses.Add(addr); }
                    }
                }
            }
        }

        var changedAddressNames = changedAddresses
            .Select(a =>
            {
                var fullName = $"{a.Vorname} {a.Nachname}".Trim();
                if (!string.IsNullOrWhiteSpace(fullName)) { return $"• {fullName}"; }
                if (!string.IsNullOrWhiteSpace(a.Unternehmen)) { return $"• {a.Unternehmen}"; }
                return "• [N. n.]";
            })
            .OrderBy(n => n)
            .ToList();

        var addressChangesCount = changedAddressNames.Count;
        var otherChanges = allRealChanges.Where(e => e.Metadata.ClrType != typeof(Adresse)).ToList();
        var otherChangesCount = otherChanges.Count;

        // KORREKTUR: "AdresseGruppen" über den Namen abfragen
        var groupCount = otherChanges.Count(e =>
            e.Metadata.ClrType == typeof(Gruppe) ||
            e.Metadata.Name == "AdresseGruppen");

        var photoCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Foto));
        var docCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Dokument));

        //// KORREKTUR: Liste vorziehen, damit sie in beiden if-Zweigen verfügbar ist
        //var detailsList = new List<string>();
        //if (groupCount > 0) { detailsList.Add(groupCount == 1 ? "einer Gruppenzuordnung" : $"{groupCount} Gruppenzuordnungen"); }
        //if (photoCount > 0) { detailsList.Add(photoCount == 1 ? "einem Foto" : $"{photoCount} Fotos"); }
        //if (docCount > 0) { detailsList.Add(docCount == 1 ? "einem Dokument" : $"{docCount} Dokumenten"); }

        //var heading = "Änderungen speichern?";
        //var text = string.Empty;

        //if (addressChangesCount > 0)
        //{
        //    heading = addressChangesCount == 1 ? "Möchten Sie die Änderung speichern?" : "Möchten Sie die Änderungen speichern?";
        //    text = addressChangesCount > 10
        //        ? string.Join(Environment.NewLine, changedAddressNames.Take(10)) + Environment.NewLine + "…"
        //        : string.Join(Environment.NewLine, changedAddressNames);

        //    // KORREKTUR: Den Zusatzhinweis mit in den Adress-Block aufnehmen
        //    if (detailsList.Count > 0)
        //    {
        //        text += $"{Environment.NewLine}{Environment.NewLine}Zusätzlich gab es Änderungen an: {string.Join(", ", detailsList)}.";
        //    }
        //}
        //else
        //{
        //    heading = otherChangesCount == 1 ? "Änderung an Zusatzdaten speichern?" : $"Änderungen an {otherChangesCount} Zusatzdaten speichern?";

        //    if (detailsList.Count > 0) { text = "Es wurden Änderungen an " + string.Join(", ", detailsList) + " vorgenommen."; }
        //    else { text = "Es wurden Änderungen an Zusatzdaten vorgenommen."; }
        //}

        //return new ChangeAnalysisResult(true, allRealChanges, heading, text);

        // Für den Expander nutzen wir direkte Zahlen statt grammatikalischer Beugung
        var detailsList = new List<string>();
        if (groupCount > 0) { detailsList.Add(groupCount == 1 ? "1 Gruppenzuordnung" : $"{groupCount} Gruppenzuordnungen"); }
        if (photoCount > 0) { detailsList.Add(photoCount == 1 ? "1 Foto" : $"{photoCount} Fotos"); }
        if (docCount > 0) { detailsList.Add(docCount == 1 ? "1 Dokument" : $"{docCount} Dokumente"); }

        var heading = string.Empty;
        var text = string.Empty;
        var expanderText = string.Empty;

        if (addressChangesCount > 0)
        {
            heading = addressChangesCount == 1 ? "Möchten Sie die Änderung speichern?" : "Möchten Sie die Änderungen speichern?";
            text = addressChangesCount == 1 ? "Es wurde eine Adresse geändert." : $"Es wurden {addressChangesCount} Adressen geändert.";

            //if (detailsList.Count > 0)
            //{
            //    text += $"{Environment.NewLine}Zusätzlich gibt es Änderungen an verknüpften Zusatzdaten.";
            //}

            // Expander füllen: Erst die Adressen, dann (falls vorhanden) die Zusatzdaten
            var expanderLines = new List<string>(changedAddressNames);
            if (detailsList.Count > 0)
            {
                expanderLines.Add(string.Empty);
                expanderLines.Add("Zusatzdaten:");
                expanderLines.AddRange(detailsList.Select(d => $"• {d}"));
            }
            expanderText = string.Join(Environment.NewLine, expanderLines);
        }
        else
        {
            heading = otherChangesCount == 1 ? "Änderung an Zusatzdaten speichern?" : $"Änderungen an {otherChangesCount} Zusatzdaten speichern?";
            text = "Es wurden ausschließlich Änderungen an verknüpften Elementen (z. B. Gruppen, Fotos) vorgenommen.";

            // Expander füllen: Nur die Zusatzdaten
            expanderText = string.Join(Environment.NewLine, detailsList.Select(d => $"• {d}"));
        }

        return new ChangeAnalysisResult(true, allRealChanges, heading, text, expanderText);
    }

    public static async Task RevertChangesAsync(List<EntityEntry> entries)
    {
        foreach (var entry in entries)
        {
            switch (entry.State)
            {
                case EntityState.Modified:
                case EntityState.Deleted:
                    await entry.ReloadAsync().ConfigureAwait(false);  // Reload setzt Modified auf Unchanged zurück und lädt alte Werte
                    break;
                case EntityState.Added:
                    entry.State = EntityState.Detached;
                    break;
            }
        }
    }

    public static bool IsEntryReallyChanged(EntityEntry entry)
    {
        if (entry.State == EntityState.Added || entry.State == EntityState.Deleted) { return true; }
        if (entry.State != EntityState.Modified) { return false; }
        if (!entry.Properties.Any(p => p.IsModified)) { return true; }
        foreach (var prop in entry.Properties)
        {
            if (!prop.IsModified) { continue; }
            var current = prop.CurrentValue;
            var original = prop.OriginalValue;
            if (Equals(original, current)) { continue; }  // direkter Vergleich (für Zahlen, Datum, etc.)
            if (prop.Metadata.ClrType == typeof(string))  // Spezialbehandlung für Strings (null == empty, trim)
            {
                var sOriginal = (original as string ?? string.Empty).Trim();
                var sCurrent = (current as string ?? string.Empty).Trim();

                if (sOriginal == sCurrent) { continue; }
            }
            return true;  // Wenn wir hier ankommen, gab es eine echte Änderung in einem Property
        }
        return false;  // Wenn alle "Modified" Properties eigentlich nur Whitespace-Unterschiede waren, oder wenn EntityState.Modified gesetzt wurde, aber keine Werte anders sind:
    }
}
