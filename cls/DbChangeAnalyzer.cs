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
        var allRealChanges = context.ChangeTracker.Entries().Where(IsEntryReallyChanged).ToList();
        //var allRealChanges = context.ChangeTracker.Entries().Where(IsEntryReallyChanged).Where(e => e.Metadata.ClrType == typeof(Adresse)).ToList();  // Nur die Hauptobjekte für die Zählung
        if (allRealChanges.Count == 0) { return new ChangeAnalysisResult(false, [], string.Empty, string.Empty); }
        // --- Text-Generierung ---
        var heading = string.Empty;
        var text = string.Empty;

        // Adress-Änderungen filtern und formatieren
        var changedAddressNames = allRealChanges
            .Where(e => e.Metadata.ClrType == typeof(Adresse))
            .Select(e => (Adresse)e.Entity)
            .Select(a =>
            {
                var fullName = $"{a.Vorname} {a.Nachname}".Trim();
                if (!string.IsNullOrWhiteSpace(fullName)) { return $"• {fullName}"; }
                if (!string.IsNullOrWhiteSpace(a.Unternehmen)) { return $"• {a.Unternehmen}"; }
                return "• [N. n.]";
            })
            .Distinct().ToList();
        var addressChangesCount = changedAddressNames.Count;

        // Andere Entitäten filtern
        var otherChanges = allRealChanges.Where(e => e.Metadata.ClrType != typeof(Adresse)).ToList();
        var otherChangesCount = otherChanges.Count;

        var groupCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Gruppe));
        var photoCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Foto));
        var docCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Dokument));

        if (addressChangesCount > 0)
        {
            heading = addressChangesCount == 1 ? "Möchten Sie die Änderung speichern?" : "Möchten Sie die Änderungen speichern?";

            if (addressChangesCount > 10) { text = string.Join(Environment.NewLine, changedAddressNames.Take(10)) + Environment.NewLine + "…"; }
            else { text = string.Join(Environment.NewLine, changedAddressNames); }

            // Zusatzhinweis
            if (otherChangesCount > 0) { text += $"{Environment.NewLine}und {otherChangesCount} Änderungen an Zusatzdaten"; }
        }
        else
        {
            heading = otherChangesCount == 1 ? "Änderung an Zusatzdaten speichern?" : $"Änderungen an {otherChangesCount} Zusatzdaten speichern?";

            var detailsList = new List<string>();
            if (groupCount == 1) { detailsList.Add("einer Gruppe"); }
            if (groupCount > 1) { detailsList.Add($"{groupCount} Gruppen"); }
            if (photoCount == 1) { detailsList.Add("einem Foto"); }
            if (photoCount > 1) { detailsList.Add($"{photoCount} Fotos"); }
            if (docCount == 1) { detailsList.Add("einem Dokument"); }
            if (docCount > 1) { detailsList.Add($"{docCount} Dokumenten"); }

            // Fallback für unbekannte Typen
            var remainder = otherChangesCount - (groupCount + photoCount + docCount);
            if (remainder == 1) { detailsList.Add("einem sonstigen Element"); }
            if (remainder > 1) { detailsList.Add($"{remainder} sonstigen Elementen"); }

            if (detailsList.Count > 0) { text = "Es wurden Änderungen an " + string.Join(", ", detailsList) + " vorgenommen."; }
            else { text = "Es wurden Änderungen an Zusatzdaten vorgenommen."; }
        }
        return new ChangeAnalysisResult(true, allRealChanges, heading, text);
    }

    public static async Task RevertChangesAsync(List<EntityEntry> entries)  // Verwirft alle lokalen Änderungen (Reload für Modified/Deleted, Detach für Added).
    {
        foreach (var entry in entries)
        {
            switch (entry.State)
            {
                case EntityState.Modified:
                case EntityState.Deleted:
                    await entry.ReloadAsync().ConfigureAwait(false);
                    break;
                case EntityState.Added:
                    entry.State = EntityState.Detached;
                    break;
            }
        }
    }

    public static bool IsEntryReallyChanged(EntityEntry entry)  // Prüft, ob ein Eintrag wirklich geändert wurde (ignoriert z.B. null vs empty string).
    {
        if (entry.State == EntityState.Added || entry.State == EntityState.Deleted) { return true; }
        if (entry.State != EntityState.Modified) { return false; }

        foreach (var prop in entry.Properties)
        {
            if (!prop.IsModified) { continue; }
            var current = prop.CurrentValue;
            var original = prop.OriginalValue;

            // 1. Direkter Vergleich
            if (Equals(original, current)) { continue; }

            // 2. Spezialbehandlung für Strings
            if (prop.Metadata.ClrType == typeof(string))
            {
                var sOriginal = original as string;
                var sCurrent = current as string;
                var sOrigClean = sOriginal ?? string.Empty;
                var sCurrClean = sCurrent ?? string.Empty;

                if (sOrigClean == sCurrClean) { continue; }
            }
            return true;
        }
        return false;
    }
}