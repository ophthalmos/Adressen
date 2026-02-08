using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Globalization;
using System.Text;
using System.Text.Json;
using Microsoft.EntityFrameworkCore;

namespace Adressen.cls;

internal record LegacyRawData(int Id, string? Gruppen, string? Dokumente, string? Geburtstag);

public static class DatabaseMigrator
{

    public static int GetDatabaseVersion(string filePath)
    {
        if (!File.Exists(filePath)) { return AppSettings.DatabaseSchemaVersion; } // Neue Datei wird eh mit aktuellem Schema erstellt
        var connectionString = $"Data Source={filePath};Mode=ReadOnly;";
        try
        {
            using var connection = new Microsoft.Data.Sqlite.SqliteConnection(connectionString);
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = "PRAGMA user_version;";
            return Convert.ToInt32(command.ExecuteScalar());
        }
        catch { return 0; }  // Im Fehlerfall (z.B. Datei gelockt) lieber von 0 ausgehen
    }

    /// <summary>
    /// Führt alle notwendigen Datenbank-Updates durch und setzt die Schema-Version.
    /// </summary>
    internal static bool MigrateLegacyData(AdressenDbContext context, IntPtr ownerHandle)
    {
        if (context == null)
        {
            return false;
        }

        var changesMade = false;

        try
        {
            // 1. Aktuelle Spalten ermitteln
            var dbColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var command = context.Database.GetDbConnection().CreateCommand())
            {
                command.CommandText = "SELECT name FROM pragma_table_info('Adressen')";
                context.Database.OpenConnection();
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    dbColumns.Add(reader.GetString(0));
                }
            }

            // --- SCHRITT A: Spalten umbenennen ---
            if (dbColumns.Contains("Firma"))
            {
                if (dbColumns.Contains("Unternehmen"))
                {
                    context.Database.ExecuteSql($"UPDATE Adressen SET Unternehmen = Firma WHERE (Unternehmen IS NULL OR Unternehmen = '') AND (Firma IS NOT NULL AND Firma <> '')");
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

            // Weitere Umbenennungen (Grussformel, Strasse, Praefix)
            var renames = new Dictionary<string, string>
            {
                { "Grußformel", "Grussformel" },
                { "Straße", "Strasse" },
                { "Präfix", "Praefix" }
            };

            foreach (var rename in renames)
            {
                if (dbColumns.Contains(rename.Key))
                {
#pragma warning disable EF1000, EF1002
                    context.Database.ExecuteSqlRaw($"ALTER TABLE Adressen RENAME COLUMN \"{rename.Key}\" TO {rename.Value}");
#pragma warning restore EF1000, EF1002
                    dbColumns.Remove(rename.Key);
                    dbColumns.Add(rename.Value);
                    changesMade = true;
                }
            }

            // --- SCHRITT B: Fehlende Spalten ergänzen ---
            var entityProperties = typeof(Adresse).GetProperties()
                .Where(p => p.Name != "Id"
                         && !Attribute.IsDefined(p, typeof(NotMappedAttribute))
                         && !p.GetAccessors().Any(x => x.IsVirtual));

            foreach (var prop in entityProperties)
            {
                if (!dbColumns.Contains(prop.Name))
                {
#pragma warning disable EF1000, EF1002
                    context.Database.ExecuteSqlRaw($"ALTER TABLE Adressen ADD COLUMN \"{prop.Name}\" TEXT");
#pragma warning restore EF1000, EF1002
                    changesMade = true;
                }
            }

            // --- SCHRITT B.2: Umstellung auf NOCASE (Case-Insensitivity) ---
            var needsNocaseUpdate = false;
            using (var command = context.Database.GetDbConnection().CreateCommand())
            {
                command.CommandText = "SELECT sql FROM sqlite_master WHERE type='table' AND name='Adressen'";
                context.Database.OpenConnection();
                var tableSql = command.ExecuteScalar()?.ToString() ?? "";

                // Prüfen, ob der Nachname bereits COLLATE NOCASE hat
                if (!tableSql.Contains("Nachname\" TEXT COLLATE NOCASE", StringComparison.OrdinalIgnoreCase))
                {
                    needsNocaseUpdate = true;
                }
            }

            if (needsNocaseUpdate)
            {
                context.Database.ExecuteSqlRaw("PRAGMA foreign_keys = OFF;");
                using var transaction = context.Database.BeginTransaction();
                try
                {
                    // 1. Bestehende Tabelle umbenennen
                    context.Database.ExecuteSqlRaw("ALTER TABLE Adressen RENAME TO Adressen_Old;");

                    // 2. Neue Tabelle mit korrekter COLLATE NOCASE Definition erstellen
                    // Hier habe ich alle deine Spalten aus der Create-Anweisung übernommen
                    context.Database.ExecuteSqlRaw(@"
            CREATE TABLE ""Adressen"" (
                ""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                ""Anrede"" TEXT,
                ""Praefix"" TEXT,
                ""Nachname"" TEXT COLLATE NOCASE,
                ""Vorname"" TEXT COLLATE NOCASE,
                ""Zwischenname"" TEXT,
                ""Nickname"" TEXT,
                ""Suffix"" TEXT,
                ""Unternehmen"" TEXT COLLATE NOCASE,
                ""Position"" TEXT,
                ""Strasse"" TEXT COLLATE NOCASE,
                ""PLZ"" TEXT,
                ""Ort"" TEXT COLLATE NOCASE,
                ""Postfach"" TEXT,
                ""Land"" TEXT,
                ""Betreff"" TEXT,
                ""Grussformel"" TEXT,
                ""Schlussformel"" TEXT,
                ""Geburtstag"" TEXT,
                ""Mail1"" TEXT,
                ""Mail2"" TEXT,
                ""Telefon1"" TEXT,
                ""Telefon2"" TEXT,
                ""Mobil"" TEXT,
                ""Fax"" TEXT,
                ""Internet"" TEXT,
                ""Notizen"" TEXT
            );");

                    // 3. Daten von Alt nach Neu kopieren
                    context.Database.ExecuteSqlRaw(@"
            INSERT INTO Adressen (
                Id, Anrede, Praefix, Nachname, Vorname, Zwischenname, Nickname, Suffix, 
                Unternehmen, Position, Strasse, PLZ, Ort, Postfach, Land, Betreff, 
                Grussformel, Schlussformel, Geburtstag, Mail1, Mail2, Telefon1, 
                Telefon2, Mobil, Fax, Internet, Notizen
            )
            SELECT 
                Id, Anrede, Praefix, Nachname, Vorname, Zwischenname, Nickname, Suffix, 
                Unternehmen, Position, Strasse, PLZ, Ort, Postfach, Land, Betreff, 
                Grussformel, Schlussformel, Geburtstag, Mail1, Mail2, Telefon1, 
                Telefon2, Mobil, Fax, Internet, Notizen
            FROM Adressen_Old;");

                    // 4. Alte Tabelle löschen
                    context.Database.ExecuteSqlRaw("DROP TABLE Adressen_Old;");

                    transaction.Commit();
                    changesMade = true;
                }
                catch (Exception)
                {
                    transaction.Rollback();
                    throw;
                }
                finally
                {
                    context.Database.ExecuteSqlRaw("PRAGMA foreign_keys = ON;");
                }
            }

            // --- SCHRITT C: Tabellenstruktur sicherstellen ---
            context.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""Gruppen"" (""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, ""Name"" TEXT NOT NULL);");
            context.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""AdresseGruppen"" (""AdressenId"" INTEGER NOT NULL, ""GruppenId"" INTEGER NOT NULL, PRIMARY KEY(""AdressenId"", ""GruppenId""), FOREIGN KEY(""AdressenId"") REFERENCES ""Adressen""(""Id"") ON DELETE CASCADE, FOREIGN KEY(""GruppenId"") REFERENCES ""Gruppen""(""Id"") ON DELETE CASCADE);");
            context.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""Dokumente"" (""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, ""Dateipfad"" TEXT NOT NULL, ""AdressId"" INTEGER NOT NULL, FOREIGN KEY(""AdressId"") REFERENCES ""Adressen""(""Id"") ON DELETE CASCADE);");

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

                    if (dataChanged)
                    {
                        context.Entry(adresse).State = EntityState.Modified;
                    }
                }

                context.SaveChanges();
                if (hasOldGruppen)
                {
                    context.Database.ExecuteSqlRaw("ALTER TABLE Adressen DROP COLUMN Gruppen");
                }

                if (hasOldDokumente)
                {
                    context.Database.ExecuteSqlRaw("ALTER TABLE Adressen DROP COLUMN Dokumente");
                }

                changesMade = true;
            }

            // --- SCHRITT E: Schema-Version setzen ---
            // Wir setzen die Version IMMER am Ende der Migration auf den aktuellen Stand aus den AppSettings.
            context.Database.ExecuteSqlRaw($"PRAGMA user_version = {AppSettings.DatabaseSchemaVersion}");

            if (changesMade)
            {
                context.Database.ExecuteSqlRaw("VACUUM;");
                Utils.MsgTaskDlg(ownerHandle, "Migration erfolgreich", "Die Datenbank wurde auf das neue Format aktualisiert.", TaskDialogIcon.ShieldSuccessGreenBar);
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(ownerHandle, ex);
            return false;
        }
    }
}