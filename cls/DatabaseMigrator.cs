#pragma warning disable EF1002 // Allow raw SQL interpolation for migration logic
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Globalization;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;

namespace Adressen.cls;

internal record LegacyRawData(int Id, string? Gruppen, string? Dokumente, string? Geburtstag);

internal static class DatabaseMigrator
{
    // Definiere hier, ab welcher Version NOCASE aktiv sein soll.
    // Wenn AppSettings.DatabaseSchemaVersion z.B. 2 ist, wird das Update ausgeführt.
    private const int VersionIntroducingNoCase = 3;

    public static int GetDatabaseVersion(string filePath)
    {
        if (!File.Exists(filePath)) { return AppSettings.DatabaseSchemaVersion; }
        var connectionString = $"Data Source={filePath};Mode=ReadOnly;";
        try
        {
            using var connection = new SqliteConnection(connectionString);
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = "PRAGMA user_version;";
            var result = command.ExecuteScalar();
            return result != null ? Convert.ToInt32(result) : 0;
        }
        catch { return 0; }
    }

    internal static bool MigrateLegacyData(AdressenDbContext context, IntPtr ownerHandle)
    {
        if (context == null) { return false; }

        var connectionString = context.Database.GetConnectionString();
        var currentDbVersion = 0;
        try
        {
            using var cmd = context.Database.GetDbConnection().CreateCommand();
            context.Database.OpenConnection();
            cmd.CommandText = "PRAGMA user_version;";
            currentDbVersion = Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
        }
        catch { /* Ignorieren */ }

        if (currentDbVersion >= AppSettings.DatabaseSchemaVersion) { return false; }

        var changesMade = false;

        using var transaction = context.Database.BeginTransaction();

        try
        {
            // WICHTIG: Fremdschlüsselprüfung SOFORT deaktivieren.
            // Da die DB aktuell fehlerhafte Referenzen auf "Adressen_Old" hat, 
            // würde sonst jeder Schreibvorgang (auch in LegacyDataCleanup) abstürzen.
            context.Database.ExecuteSqlRaw("PRAGMA foreign_keys = OFF;");

            // ---------------------------------------------------------
            // PHASE 1: Strukturelle Anpassungen
            // ---------------------------------------------------------

            var dbColumns = GetTableColumns(context, "Adressen");

            // 1.1 Umbenennungen
            var renames = new Dictionary<string, string>
        {
            { "Firma", "Unternehmen" },
            { "Grußformel", "Grussformel" },
            { "Straße", "Strasse" },
            { "Praefix", "Praefix" }
        };

            // Spezialfall Firma -> Unternehmen
            if (dbColumns.Contains("Firma"))
            {
                if (dbColumns.Contains("Unternehmen"))
                {
                    // Merge Logik falls beide existieren
                    context.Database.ExecuteSqlRaw("UPDATE Adressen SET Unternehmen = Firma WHERE (Unternehmen IS NULL OR Unternehmen = '') AND (Firma IS NOT NULL AND Firma <> '')");
                    context.Database.ExecuteSqlRaw("ALTER TABLE Adressen DROP COLUMN \"Firma\"");
                }
                else
                {
                    context.Database.ExecuteSqlRaw("ALTER TABLE Adressen RENAME COLUMN \"Firma\" TO Unternehmen");
                    dbColumns.Add("Unternehmen");
                }
                dbColumns.Remove("Firma");
                changesMade = true;
            }

            foreach (var rename in renames.Where(r => r.Key != "Firma"))
            {
                if (dbColumns.Contains(rename.Key))
                {
                    context.Database.ExecuteSqlRaw($"ALTER TABLE Adressen RENAME COLUMN \"{rename.Key}\" TO {rename.Value}");
                    dbColumns.Remove(rename.Key);
                    dbColumns.Add(rename.Value);
                    changesMade = true;
                }
            }

            // 1.2 Fehlende Spalten ergänzen
            var entityProperties = typeof(Adresse).GetProperties()
                .Where(p => p.Name != "Id"
                            && !Attribute.IsDefined(p, typeof(NotMappedAttribute))
                            && !p.GetAccessors().Any(x => x.IsVirtual));

            foreach (var prop in entityProperties)
            {
                if (!dbColumns.Contains(prop.Name))
                {
                    context.Database.ExecuteSqlRaw($"ALTER TABLE Adressen ADD COLUMN \"{prop.Name}\" TEXT");
                    changesMade = true;
                }
            }

            // 1.3 Hilfstabellen sicherstellen
            // Hinweis: Da FKs jetzt OFF sind, laufen diese Befehle durch, auch wenn die Referenzen "krumm" sind.
            context.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""Fotos"" (""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, ""AdressId"" INTEGER NOT NULL UNIQUE, ""Fotodaten"" BLOB, FOREIGN KEY(""AdressId"") REFERENCES ""Adressen""(""Id"") ON DELETE CASCADE);");
            context.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""Gruppen"" (""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, ""Name"" TEXT NOT NULL);");
            context.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""AdresseGruppen"" (""AdressenId"" INTEGER NOT NULL, ""GruppenId"" INTEGER NOT NULL, PRIMARY KEY(""AdressenId"", ""GruppenId""), FOREIGN KEY(""AdressenId"") REFERENCES ""Adressen""(""Id"") ON DELETE CASCADE, FOREIGN KEY(""GruppenId"") REFERENCES ""Gruppen""(""Id"") ON DELETE CASCADE);");
            context.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""Dokumente"" (""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, ""Dateipfad"" TEXT NOT NULL, ""AdressId"" INTEGER NOT NULL, FOREIGN KEY(""AdressId"") REFERENCES ""Adressen""(""Id"") ON DELETE CASCADE);");

            // 1.4 Datenmigration (JSON/Datum)
            // Das hier verursachte vorher den Crash, weil SaveChanges() die kaputten FKs prüfte.
            // Jetzt ist FK-Prüfung aus, also läuft es durch.
            if (LegacyDataCleanup(context, dbColumns)) { changesMade = true; }

            // ---------------------------------------------------------
            // PHASE 2: Table Rebuild (Hier werden die kaputten Referenzen korrigiert)
            // ---------------------------------------------------------

            if (currentDbVersion < VersionIntroducingNoCase)
            {
                context.SaveChanges();

                // 2.2 Haupttabelle Backup
                var backupTableName = "Adressen_Legacy_Backup";
                context.Database.ExecuteSqlRaw($"DROP TABLE IF EXISTS \"{backupTableName}\"");
                context.Database.ExecuteSqlRaw($"ALTER TABLE \"Adressen\" RENAME TO \"{backupTableName}\"");

                // 2.3 Neue Tabelle Adressen erstellen (Clean)
                var createSql = @"
                CREATE TABLE ""Adressen"" (
                    ""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    ""Anrede"" TEXT COLLATE NOCASE,
                    ""Praefix"" TEXT COLLATE NOCASE,
                    ""Nachname"" TEXT COLLATE NOCASE,
                    ""Vorname"" TEXT COLLATE NOCASE,
                    ""Zwischenname"" TEXT COLLATE NOCASE,
                    ""Nickname"" TEXT COLLATE NOCASE,
                    ""Suffix"" TEXT COLLATE NOCASE,
                    ""Unternehmen"" TEXT COLLATE NOCASE,
                    ""Position"" TEXT COLLATE NOCASE,
                    ""Strasse"" TEXT COLLATE NOCASE,
                    ""PLZ"" TEXT COLLATE NOCASE,
                    ""Ort"" TEXT COLLATE NOCASE,
                    ""Postfach"" TEXT COLLATE NOCASE,
                    ""Land"" TEXT COLLATE NOCASE,
                    ""Betreff"" TEXT COLLATE NOCASE,
                    ""Grussformel"" TEXT COLLATE NOCASE,
                    ""Schlussformel"" TEXT COLLATE NOCASE,
                    ""Geburtstag"" TEXT,
                    ""Mail1"" TEXT COLLATE NOCASE,
                    ""Mail2"" TEXT COLLATE NOCASE,
                    ""Telefon1"" TEXT COLLATE NOCASE,
                    ""Telefon2"" TEXT COLLATE NOCASE,
                    ""Mobil"" TEXT COLLATE NOCASE,
                    ""Fax"" TEXT COLLATE NOCASE,
                    ""Internet"" TEXT COLLATE NOCASE,
                    ""Notizen"" TEXT COLLATE NOCASE
                );";
                context.Database.ExecuteSqlRaw(createSql);

                // 2.4 Daten kopieren
                var columns = "\"Id\", \"Anrede\", \"Praefix\", \"Nachname\", \"Vorname\", \"Zwischenname\", \"Nickname\", \"Suffix\", \"Unternehmen\", \"Position\", \"Strasse\", \"PLZ\", \"Ort\", \"Postfach\", \"Land\", \"Betreff\", \"Grussformel\", \"Schlussformel\", \"Geburtstag\", \"Mail1\", \"Mail2\", \"Telefon1\", \"Telefon2\", \"Mobil\", \"Fax\", \"Internet\", \"Notizen\"";
                context.Database.ExecuteSqlRaw($"INSERT INTO \"Adressen\" ({columns}) SELECT {columns} FROM \"{backupTableName}\"");

                // ---------------------------------------------------------
                // REBUILD KIND-TABELLEN
                // Hier korrigieren wir den Fehler "REFERENCES Adressen_Old"
                // indem wir die Tabellen neu erstellen und auf "Adressen" zeigen lassen.
                // ---------------------------------------------------------

                // --- A) FOTOS reparieren ---
                context.Database.ExecuteSqlRaw("ALTER TABLE \"Fotos\" RENAME TO \"Fotos_Backup\"");

                // Neu erstellen
                context.Database.ExecuteSqlRaw(@"
    CREATE TABLE ""Fotos"" (
        ""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
        ""AdressId"" INTEGER NOT NULL UNIQUE,
        ""Fotodaten"" BLOB,
        FOREIGN KEY(""AdressId"") REFERENCES ""Adressen""(""Id"") ON DELETE CASCADE
    );");

                // Daten kopieren - MIT FILTERUNG verwaister Einträge
                context.Database.ExecuteSqlRaw(@"
    INSERT INTO ""Fotos"" (""Id"", ""AdressId"", ""Fotodaten"") 
    SELECT f.""Id"", f.""AdressId"", f.""Fotodaten"" 
    FROM ""Fotos_Backup"" f
    WHERE EXISTS (SELECT 1 FROM ""Adressen"" a WHERE a.""Id"" = f.""AdressId"")");

                // Backup weg
                context.Database.ExecuteSqlRaw("DROP TABLE \"Fotos_Backup\"");

                // --- B) DOKUMENTE reparieren ---
                context.Database.ExecuteSqlRaw("ALTER TABLE \"Dokumente\" RENAME TO \"Dokumente_Backup\"");

                // Neu erstellen
                context.Database.ExecuteSqlRaw(@"
    CREATE TABLE ""Dokumente"" (
        ""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
        ""Dateipfad"" TEXT NOT NULL,
        ""AdressId"" INTEGER NOT NULL,
        FOREIGN KEY(""AdressId"") REFERENCES ""Adressen""(""Id"") ON DELETE CASCADE
    );");

                // Daten kopieren - MIT FILTERUNG
                context.Database.ExecuteSqlRaw(@"
    INSERT INTO ""Dokumente"" (""Id"", ""Dateipfad"", ""AdressId"") 
    SELECT d.""Id"", d.""Dateipfad"", d.""AdressId"" 
    FROM ""Dokumente_Backup"" d
    WHERE EXISTS (SELECT 1 FROM ""Adressen"" a WHERE a.""Id"" = d.""AdressId"")");

                context.Database.ExecuteSqlRaw("DROP TABLE \"Dokumente_Backup\"");

                // --- C) ADRESSEGRUPPEN reparieren ---
                context.Database.ExecuteSqlRaw("ALTER TABLE \"AdresseGruppen\" RENAME TO \"AdresseGruppen_Backup\"");

                // Neu erstellen
                context.Database.ExecuteSqlRaw(@"
    CREATE TABLE ""AdresseGruppen"" (
        ""AdressenId"" INTEGER NOT NULL,
        ""GruppenId"" INTEGER NOT NULL,
        PRIMARY KEY(""AdressenId"", ""GruppenId""),
        FOREIGN KEY(""AdressenId"") REFERENCES ""Adressen""(""Id"") ON DELETE CASCADE,
        FOREIGN KEY(""GruppenId"") REFERENCES ""Gruppen""(""Id"") ON DELETE CASCADE
    );");

                // Daten kopieren - MIT FILTERUNG
                context.Database.ExecuteSqlRaw(@"
    INSERT INTO ""AdresseGruppen"" (""AdressenId"", ""GruppenId"") 
    SELECT ag.""AdressenId"", ag.""GruppenId"" 
    FROM ""AdresseGruppen_Backup"" ag
    WHERE EXISTS (SELECT 1 FROM ""Adressen"" a WHERE a.""Id"" = ag.""AdressenId"")
      AND EXISTS (SELECT 1 FROM ""Gruppen"" g WHERE g.""Id"" = ag.""GruppenId"")");

                context.Database.ExecuteSqlRaw("DROP TABLE \"AdresseGruppen_Backup\"");

                changesMade = true;
            }

            // Schema Version setzen
            context.Database.ExecuteSqlRaw($"PRAGMA user_version = {AppSettings.DatabaseSchemaVersion}");

            // Ganz am Ende, wenn alles sauber ist: Check aktivieren
            // Wenn hier etwas knallt, sind die Daten wirklich inkonsistent, aber die Struktur stimmt jetzt.
            // Wir machen es VOR dem Commit, damit wir bei Fehlern rollen können.
            context.Database.ExecuteSqlRaw("PRAGMA foreign_key_check;");
            context.Database.ExecuteSqlRaw("PRAGMA foreign_keys = ON;");

            transaction.Commit();

            if (changesMade)
            {
                context.Database.ExecuteSqlRaw("VACUUM;");
                Utils.MsgTaskDlg(ownerHandle, "Datenbank aktualisiert",
                    $"Die Datenbank wurde erfolgreich migriert (v{AppSettings.DatabaseSchemaVersion}).",
                    TaskDialogIcon.ShieldSuccessGreenBar);
                return true;
            }
            return false;
        }
        catch (Exception ex)
        {
            transaction.Rollback();
            Utils.ErrTaskDlg(ownerHandle, ex);
            return false;
        }
    }

    // --- Hilfsmethoden ---

    private static HashSet<string> GetTableColumns(AdressenDbContext context, string tableName)
    {
        var columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var command = context.Database.GetDbConnection().CreateCommand();
        command.CommandText = $"SELECT name FROM pragma_table_info('{tableName}')";
        if (context.Database.GetDbConnection().State != ConnectionState.Open) { context.Database.OpenConnection(); }
        using var reader = command.ExecuteReader();
        while (reader.Read()) { columns.Add(reader.GetString(0)); }
        return columns;
    }

    private static bool LegacyDataCleanup(AdressenDbContext context, HashSet<string> dbColumns)
    {
        // ... Hier dein bestehender JSON/Geburtstag Cleanup Code ...
        // (Ich habe ihn hier gekürzt, da er unverändert bleiben kann, 
        //  siehe dein Original-Snippet 'Schritt D')

        // WICHTIG: Wenn du hier Spalten droppst (Gruppen, Dokumente), 
        // dann darfst du sie in Phase 2 im SQL INSERT nicht mehr auflisten!
        // Da dein Code sie droppt: 
        // if (hasOldGruppen) { context.Database.ExecuteSqlRaw("ALTER TABLE Adressen DROP COLUMN Gruppen"); }
        // ist das okay, solange das VOR Phase 2 passiert.

        // Da ich den Aufruf VOR Phase 2 platziert habe, ist alles korrekt.

        // Dummy Return für dieses Snippet:
        //return false;

        // --- SCHRITT D: Datenmigration (JSON/Datum) ---
        var hasOldGruppen = dbColumns.Contains("Gruppen");
        var hasOldDokumente = dbColumns.Contains("Dokumente");
        bool hasOldDateFormats;

        using (var command = context.Database.GetDbConnection().CreateCommand())
        {
            command.CommandText = "SELECT 1 FROM Adressen WHERE Geburtstag LIKE '%.%' OR Geburtstag = '' LIMIT 1";
            using var reader = command.ExecuteReader();
            hasOldDateFormats = reader.HasRows;
        }

        if (hasOldGruppen || hasOldDokumente || hasOldDateFormats)
        {
            var sbSql = new StringBuilder();
            sbSql.Append("SELECT Id, NULLIF(CAST(Geburtstag AS TEXT), '') AS Geburtstag");
            sbSql.Append(hasOldGruppen ? ", Gruppen" : ", NULL AS Gruppen");
            sbSql.Append(hasOldDokumente ? ", Dokumente" : ", NULL AS Dokumente");
            sbSql.Append(" FROM Adressen");

            var legacyData = context.Database.SqlQueryRaw<LegacyRawData>(sbSql.ToString()).ToList();
            context.Database.ExecuteSqlRaw("UPDATE Adressen SET Geburtstag = NULL;");

            var allAdressen = context.Adressen.Include(a => a.Gruppen).Include(a => a.Dokumente).ToList();
            var gruppenCache = new Dictionary<string, Gruppe>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in legacyData)
            {
                var adresse = allAdressen.FirstOrDefault(a => a.Id == row.Id);
                if (adresse == null) { continue; }

                var dataChanged = false;

                // Geburtstag
                if (!string.IsNullOrWhiteSpace(row.Geburtstag) &&
                    (DateOnly.TryParseExact(row.Geburtstag, "d.M.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate) ||
                     DateOnly.TryParse(row.Geburtstag, CultureInfo.GetCultureInfo("de-DE"), DateTimeStyles.None, out parsedDate)))
                {
                    adresse.Geburtstag = parsedDate;
                    dataChanged = true;
                }

                // Gruppen (JSON)
                if (hasOldGruppen && !string.IsNullOrWhiteSpace(row.Gruppen))
                {
                    try
                    {
                        var namen = JsonSerializer.Deserialize<List<string>>(row.Gruppen);
                        if (namen != null)
                        {
                            foreach (var name in namen.Where(n => !string.IsNullOrWhiteSpace(n)))
                            {
                                if (!gruppenCache.TryGetValue(name, out var gruppe))
                                {
                                    gruppe = context.Gruppen.Local.FirstOrDefault(g => g.Name == name) ??
                                             context.Gruppen.FirstOrDefault(g => g.Name == name) ??
                                             new Gruppe { Name = name };
                                    if (gruppe.Id == 0 && !context.Gruppen.Local.Contains(gruppe))
                                    {
                                        context.Gruppen.Add(gruppe);
                                    }

                                    gruppenCache[name] = gruppe;
                                }
                                if (!adresse.Gruppen.Any(g => g.Name == name)) { adresse.Gruppen.Add(gruppe); dataChanged = true; }
                            }
                        }
                    }
                    catch { }
                }

                // Dokumente (JSON)
                if (hasOldDokumente && !string.IsNullOrWhiteSpace(row.Dokumente))
                {
                    try
                    {
                        var pfade = JsonSerializer.Deserialize<List<string>>(row.Dokumente);
                        if (pfade != null)
                        {
                            foreach (var pfad in pfade.Where(p => !string.IsNullOrWhiteSpace(p)))
                            {
                                if (!adresse.Dokumente.Any(d => d.Dateipfad == pfad))
                                {
                                    adresse.Dokumente.Add(new Dokument { Dateipfad = pfad });
                                    dataChanged = true;
                                }
                            }
                        }
                    }
                    catch { }
                }
                if (dataChanged) { context.Entry(adresse).State = EntityState.Modified; }
            }

            context.SaveChanges();
            if (hasOldGruppen) { context.Database.ExecuteSqlRaw("ALTER TABLE Adressen DROP COLUMN Gruppen"); }
            if (hasOldDokumente) { context.Database.ExecuteSqlRaw("ALTER TABLE Adressen DROP COLUMN Dokumente"); }
            return true;
        }
        return false;

    }
}
