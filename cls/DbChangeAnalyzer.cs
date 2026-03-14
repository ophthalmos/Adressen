using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;

namespace Adressen.cls;

public static class DbChangeAnalyzer
{
    //public record ChangeAnalysisResult(bool HasChanges, List<EntityEntry> RealChanges, string DialogHeading, string DialogText);

    //public static ChangeAnalysisResult AnalyzeChanges(DbContext? context)
    //{
    //    if (context == null) { return new ChangeAnalysisResult(false, [], string.Empty, string.Empty); }

    //    context.ChangeTracker.DetectChanges();
    //    var allRealChanges = context.ChangeTracker.Entries().Where(IsEntryReallyChanged).ToList();

    //    if (allRealChanges.Count == 0) { return new ChangeAnalysisResult(false, [], string.Empty, string.Empty); }

    //    var changedAddresses = allRealChanges
    //        .Where(e => e.Metadata.ClrType == typeof(Adresse))
    //        .Select(e => (Adresse)e.Entity)
    //        .ToHashSet();

    //    var shadowEntries = allRealChanges.Where(e => e.Metadata.Name == "AdresseGruppen");

    //    foreach (var shadow in shadowEntries)
    //    {
    //        foreach (var fk in shadow.Metadata.GetForeignKeys())
    //        {
    //            if (fk.PrincipalEntityType.ClrType == typeof(Adresse))
    //            {
    //                var fkProp = fk.Properties.Count > 0 ? fk.Properties[0] : null;
    //                if (fkProp != null && shadow.CurrentValues[fkProp] is int addressId)
    //                {
    //                    var addr = context.Set<Adresse>().Local.FirstOrDefault(a => a.Id == addressId);
    //                    if (addr != null) { changedAddresses.Add(addr); }
    //                }
    //            }
    //        }
    //    }

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
    //    var otherChanges = allRealChanges.Where(e => e.Metadata.ClrType != typeof(Adresse)).ToList();
    //    var otherChangesCount = otherChanges.Count;

    //    var groupCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Gruppe) || e.Metadata.Name == "AdresseGruppen");
    //    var photoCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Foto));
    //    var docCount = otherChanges.Count(e => e.Metadata.ClrType == typeof(Dokument));

    //    var detailsList = new List<string>();
    //    if (groupCount > 0) { detailsList.Add(groupCount == 1 ? "1 Gruppenzuordnung" : $"{groupCount} Gruppenzuordnungen"); }
    //    if (photoCount > 0) { detailsList.Add(photoCount == 1 ? "1 Foto" : $"{photoCount} Fotos"); }
    //    if (docCount > 0) { detailsList.Add(docCount == 1 ? "1 Dokument" : $"{docCount} Dokumente"); }

    //    var heading = string.Empty;
    //    var text = string.Empty;

    //    if (addressChangesCount > 0)
    //    {
    //        heading = addressChangesCount == 1 ? "Möchtest du die Änderung speichern?" : "Möchtest du die Änderungen speichern?";
    //        var textLines = new List<string> { "Adressen:" };
    //        if (addressChangesCount > 12)
    //        {
    //            textLines.AddRange(changedAddressNames.Take(9));
    //            textLines.Add("  …");
    //            textLines.AddRange(changedAddressNames.TakeLast(3));
    //        }
    //        else { textLines.AddRange(changedAddressNames); }

    //        if (detailsList.Count > 0)
    //        {
    //            textLines.Add(string.Empty);
    //            textLines.Add("Zusatzdaten:");
    //            textLines.AddRange(detailsList.Select(d => $"• {d}"));
    //        }
    //        text = string.Join(Environment.NewLine, textLines);
    //    }
    //    else
    //    {
    //        heading = otherChangesCount == 1 ? "Möchtest du die Änderung speichern?" : "Möchtest du die Änderungen speichern?";
    //        text = string.Join(Environment.NewLine, detailsList.Select(d => $"• {d}"));
    //    }
    //    return new ChangeAnalysisResult(true, allRealChanges, heading, text);
    //}

    public record ChangeAnalysisResult(bool HasChanges, List<EntityEntry> RealChanges, string DialogHeading, string DialogText, string ExpanderText);

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

        // 1. Gruppen-Zuordnungen auflösen (Shadow-Entities)
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

        // 2. Foto-Änderungen auflösen und der Adresse zuordnen
        var photoEntries = allRealChanges.Where(e => e.Metadata.ClrType == typeof(Foto));
        foreach (var photoEntry in photoEntries)
        {
            if (photoEntry.Entity is Foto foto)
            {
                var addr = context.Set<Adresse>().Local.FirstOrDefault(a => a.Id == foto.AdressId);
                if (addr != null) { changedAddresses.Add(addr); }
            }
        }

        // 3. Dokument-Änderungen auflösen und der Adresse zuordnen
        var docEntries = allRealChanges.Where(e => e.Metadata.ClrType == typeof(Dokument));
        foreach (var docEntry in docEntries)
        {
            if (docEntry.Entity is Dokument doc)
            {
                var addr = context.Set<Adresse>().Local.FirstOrDefault(a => a.Id == doc.AdressId);
                if (addr != null) { changedAddresses.Add(addr); }
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

        // Echte, systemweite Gruppen-Änderungen ermitteln (z.B. über FrmGroupsEdit)
        var deletedGroupsCount = allRealChanges.Count(e => e.Metadata.ClrType == typeof(Gruppe) && e.State == EntityState.Deleted);
        var otherGroupsCount = allRealChanges.Count(e => e.Metadata.ClrType == typeof(Gruppe) && (e.State == EntityState.Added || e.State == EntityState.Modified));

        var totalChangesCount = addressChangesCount + deletedGroupsCount + otherGroupsCount;
        var heading = totalChangesCount == 1 ? "Möchtest du die Änderung speichern?" : "Möchtest du die Änderungen speichern?";

        var text = string.Empty;
        var expanderLines = new List<string>();

        // Haupttext und Expander-Zeilen aufbauen
        if (addressChangesCount > 0)
        {
            text = addressChangesCount == 1 ? "Eine Adresse wurde verändert." : $"Es wurden Änderungen an {addressChangesCount} Adressen vorgenommen.";
            if (addressChangesCount > 13)  // greift erst ab 14 Einträgen; addressChangesCount - 12 ergibt mindestens 2, also "… (2 weitere)"
            {
                expanderLines.AddRange(changedAddressNames.Take(9));
                expanderLines.Add($"  … ({addressChangesCount - 12} weitere)");  // falls es genau 13 Änderungen gibt, wird hier NICHT "… (1 weitere)" angezeigt
                expanderLines.AddRange(changedAddressNames.TakeLast(3));
            }
            else { expanderLines.AddRange(changedAddressNames); }
        }

        if (deletedGroupsCount > 0 || otherGroupsCount > 0)
        {
            if (string.IsNullOrEmpty(text))
            {
                text = "Es wurden systemweite Änderungen an den Gruppen vorgenommen.";
            }

            if (expanderLines.Count > 0) { expanderLines.Add(string.Empty); }

            if (deletedGroupsCount > 0)
            {
                expanderLines.Add(deletedGroupsCount == 1 ? "Es wurde 1 Gruppe entfernt." : $"Es wurden {deletedGroupsCount} Gruppen entfernt.");
            }
            if (otherGroupsCount > 0)
            {
                expanderLines.Add(otherGroupsCount == 1 ? "Es wurde 1 Gruppe hinzugefügt/geändert." : $"Es wurden {otherGroupsCount} Gruppen hinzugefügt/geändert.");
            }
        }

        var expanderText = string.Join(Environment.NewLine, expanderLines);

        return new ChangeAnalysisResult(true, allRealChanges, heading, text, expanderText);
    }


    public static async Task RevertChangesAsync(List<EntityEntry> entries, BindingSource? addressSource = null)
    {
        foreach (var entry in entries)
        {
            switch (entry.State)
            {
                case EntityState.Modified:
                    // Setzt die Werte lokal auf den Stand beim Laden zurück - OHNE DB-Abfrage!
                    entry.CurrentValues.SetValues(entry.OriginalValues);
                    entry.State = EntityState.Unchanged;
                    break;

                case EntityState.Deleted:
                    entry.State = EntityState.Unchanged;
                    if (addressSource is not null && entry.Entity is Adresse adresse)
                    {
                        if (!addressSource.Contains(adresse)) { addressSource.Add(adresse); }
                    }
                    break;

                case EntityState.Added:
                    entry.State = EntityState.Detached;
                    break;
            }
        }
        await Task.CompletedTask;
    }

    public static bool IsEntryReallyChanged(EntityEntry entry)
    {
        if (entry.State == EntityState.Added || entry.State == EntityState.Deleted) { return true; }
        if (entry.State != EntityState.Modified) { return false; }

        foreach (var prop in entry.Properties)
        {
            if (!prop.IsModified) { continue; }

            var current = prop.CurrentValue;
            var original = prop.OriginalValue;

            if (Equals(original, current)) { continue; }

            if (prop.Metadata.ClrType == typeof(string))
            {
                var sOriginal = (original as string ?? string.Empty).Trim();
                var sCurrent = (current as string ?? string.Empty).Trim();

                if (sOriginal == sCurrent) { continue; }
            }
            return true;
        }
        return false;
    }
}