using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;

namespace Adressen.cls;

public static class DbChangeAnalyzer
{
    public record ChangeAnalysisResult(bool HasChanges, List<EntityEntry> RealChanges, string DialogHeading, string DialogText);

    public static ChangeAnalysisResult AnalyzeChanges(DbContext? context)
    {
        if (context == null) { return new ChangeAnalysisResult(false, [], string.Empty, string.Empty); }

        context.ChangeTracker.DetectChanges();

        // Wir holen alle echten Änderungen (ignoriert reine Whitespace-Änderungen)
        var allRealChanges = context.ChangeTracker.Entries().Where(IsEntryReallyChanged).ToList();

        if (allRealChanges.Count == 0) { return new ChangeAnalysisResult(false, [], string.Empty, string.Empty); }

        // --- 1. Echte Adress-Änderungen (Namen, Telefon etc.) ---
        var changedAddresses = allRealChanges
            .Where(e => e.Metadata.ClrType == typeof(Adresse))
            .Select(e => (Adresse)e.Entity)
            .ToHashSet(); // HashSet verhindert Duplikate

        // --- 2. Indirekte Adress-Änderungen (Gruppen-Zuordnungen) ---
        // Wir suchen Einträge, die keine Klasse haben (Schatten-Entitäten) und "Gruppe" im Namen tragen
        var shadowEntries = allRealChanges.Where(e => e.Metadata.ClrType == null && e.Metadata.Name.Contains("Gruppe"));

        foreach (var shadow in shadowEntries)
        {
            // Wir suchen den Fremdschlüssel, der zur Adresse zeigt
            foreach (var fk in shadow.Metadata.GetForeignKeys())
            {
                if (fk.PrincipalEntityType.ClrType == typeof(Adresse))
                {
                    // Wert des Fremdschlüssels (AdressId) aus der Schatten-Entität lesen
                    // Da es meist nur eine Property für den FK gibt (AdressId), nehmen wir die erste.
                    // Wir prüfen die Anzahl und greifen direkt auf den Index 0 zu
                    var fkProp = fk.Properties.Count > 0 ? fk.Properties[0] : null;
                    if (fkProp != null && shadow.CurrentValues[fkProp] is int addressId)
                    {
                        var addr = context.Set<Adresse>().Local.FirstOrDefault(a => a.Id == addressId);
                        if (addr != null) { changedAddresses.Add(addr); }
                    }
                }
            }
        }

        // --- Text-Generierung ---
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

        // 2. Andere Entitäten filtern
        var otherChanges = allRealChanges.Where(e => e.Metadata.ClrType != typeof(Adresse)).ToList();
        var otherChangesCount = otherChanges.Count;

        var groupCount = otherChanges.Count(e =>
            e.Metadata.ClrType == typeof(Gruppe) ||
            (e.Metadata.ClrType == null && e.Metadata.Name.Contains("Gruppe"))); // Erkennt "AdresseGruppe" Join-Table ohne Klasse

        var photoCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Foto));
        var docCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Dokument));

        var heading = addressChangesCount > 0
        ? (addressChangesCount == 1 ? "Möchten Sie die Änderung speichern?" : "Möchten Sie die Änderungen speichern?")
        : "Änderungen speichern?";

        var text = string.Empty;
        if (addressChangesCount > 0)
        {
            heading = addressChangesCount == 1 ? "Möchten Sie die Änderung speichern?" : "Möchten Sie die Änderungen speichern?";
            text = addressChangesCount > 10
            ? string.Join(Environment.NewLine, changedAddressNames.Take(10)) + Environment.NewLine + "…"
            : string.Join(Environment.NewLine, changedAddressNames);
            // Zusatzhinweis
            //if (otherChangesCount > 0) { text += $"{Environment.NewLine}und {otherChangesCount} Änderungen an Zusatzdaten"; }
        }
        else
        {
            // Wenn nur Gruppen/Fotos geändert wurden, aber keine Adress-Texte
            heading = otherChangesCount == 1 ? "Änderung an Zusatzdaten speichern?" : $"Änderungen an {otherChangesCount} Zusatzdaten speichern?";

            var detailsList = new List<string>();
            if (groupCount > 0) { detailsList.Add(groupCount == 1 ? "einer Gruppenzuordnung" : $"{groupCount} Gruppenzuordnungen"); }
            if (photoCount > 0) { detailsList.Add(photoCount == 1 ? "einem Foto" : $"{photoCount} Fotos"); }
            if (docCount > 0) { detailsList.Add(docCount == 1 ? "einem Dokument" : $"{docCount} Dokumenten"); }
            //var remainder = otherChangesCount - (groupCount + photoCount + docCount);  
            //if (remainder > 0) { detailsList.Add(remainder == 1 ? "einem sonstigen Element" : $"{remainder} sonstigen Elementen"); }

            if (detailsList.Count > 0) { text = "Es wurden Änderungen an " + string.Join(", ", detailsList) + " vorgenommen."; }
            else { text = "Es wurden Änderungen an Zusatzdaten vorgenommen."; }
        }
        return new ChangeAnalysisResult(true, allRealChanges, heading, text);
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
