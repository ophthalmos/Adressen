using System.Globalization;
using System.Text;
using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;

namespace Adressen.cls;

/// <summary>Zustand einer Datenbankdatei aus Sicht des aktuellen Schemas.</summary>
internal enum DbState
{
    Fresh,        // Datei fehlt oder enthält keine Tabellen → Schema komplett aus dem EF-Modell anlegen
    Current,      // Struktur vollständig; ein fehlender Versionsstempel wird nur nachgezogen
    NeedsUpgrade, // Struktur bekannt, aber Spalten des Modells fehlen (künftige Versionen) → modellgetrieben ergänzen
    Legacy,       // Format vor Version 3 (JSON-Gruppen, alte Spaltennamen, Tabellen ohne NOCASE) → nur mit Adressen 1.2.8 konvertierbar
    TooNew        // user_version ist größer als die des Programms
}

internal sealed record ColumnSpec(string Name, string Ddl);
internal sealed record DbInspection(DbState State, int Version, IReadOnlyList<ColumnSpec> MissingColumns);

/// <summary>
/// Schema-Verwaltung der SQLite-Datenbank. Das EF-Modell (AdressenDbContext) ist die einzige Schemaquelle;
/// <see cref="AppSettings.DatabaseSchemaVersion"/> (= 5) ist der Urzustand. Ältere Formate werden nicht mehr migriert,
/// sondern erkannt und mit Hinweis auf Adressen 1.2.8 abgewiesen. Künftige Änderungen: neue Spalte im Modell ergänzen,
/// Versionsnummer erhöhen – fehlende Spalten werden hier automatisch per ALTER TABLE ADD COLUMN nachgezogen.
/// Alles darüber hinaus (Umbenennen, Umbau) ist ein bewusster eigener Schritt.
/// </summary>
internal static class DatabaseMigrator
{
    private static readonly string[] RequiredTables = ["Adressen", "Gruppen", "Dokumente", "Fotos", "AdresseGruppen"];
    private static readonly string[] LegacyColumns = ["Firma", "Gruppen", "Dokumente", "Straße", "Grußformel", "Präfix"];  // gab es nur vor Version 3

    /// <summary>Prüft die Datei nur lesend und ohne EF-Context (die Verbindung wird danach wieder geschlossen).</summary>
    public static DbInspection Inspect(string filePath)
    {
        if (!File.Exists(filePath)) { return new DbInspection(DbState.Fresh, AppSettings.DatabaseSchemaVersion, []); }
        using var connection = new SqliteConnection($"Data Source={filePath}");
        connection.Open();
        var version = Convert.ToInt32(Scalar(connection, "PRAGMA user_version;") ?? 0);
        if (version > AppSettings.DatabaseSchemaVersion) { return new DbInspection(DbState.TooNew, version, []); }

        var tables = Names(connection, "SELECT name FROM sqlite_master WHERE type = 'table'");
        if (tables.Count == 0) { return new DbInspection(DbState.Fresh, version, []); }
        if (!RequiredTables.All(tables.Contains)) { return new DbInspection(DbState.Legacy, version, []); }

        var columns = Names(connection, "SELECT name FROM pragma_table_info('Adressen')");
        if (LegacyColumns.Any(columns.Contains)) { return new DbInspection(DbState.Legacy, version, []); }
        if (version < 3)  // vor v3 wurde die Tabelle ohne COLLATE NOCASE angelegt; nur der Versionsstempel allein ist kein Beweis
        {
            var createSql = Scalar(connection, "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'Adressen'") as string ?? string.Empty;
            if (!createSql.Contains("NOCASE", StringComparison.OrdinalIgnoreCase)) { return new DbInspection(DbState.Legacy, version, []); }
        }

        var missing = ModelColumns().Where(c => !columns.Contains(c.Name)).ToList();
        return new DbInspection(missing.Count > 0 ? DbState.NeedsUpgrade : DbState.Current, version, missing);
    }

    /// <summary>Legt eine neue Datei mit dem kompletten Schema aus dem EF-Modell an (eine vorhandene Datei wird ersetzt – der Aufrufer fragt vorher).</summary>
    public static void CreateNew(string filePath)
    {
        SqliteConnection.ClearAllPools();  // offene Pool-Verbindungen lösen, sonst scheitert das Löschen an Dateisperren
        foreach (var path in new[] { filePath, filePath + "-wal", filePath + "-shm" })  // verwaiste WAL-Dateien würden sonst in die neue Datei eingespielt
        {
            if (File.Exists(path)) { File.Delete(path); }
        }
        using var context = new AdressenDbContext(filePath);
        context.Database.EnsureCreated();
        StampVersion(context);
    }

    /// <summary>Ergänzt fehlende Modell-Spalten (Zustand NeedsUpgrade) in einer Transaktion und setzt die Schema-Version. Vorher <see cref="CreateBackupCopy"/> aufrufen.</summary>
    public static void Upgrade(AdressenDbContext context, DbInspection inspection)
    {
        using var transaction = context.Database.BeginTransaction();
        foreach (var column in inspection.MissingColumns)
        {
            var sql = $"ALTER TABLE \"Adressen\" ADD COLUMN {column.Ddl}";
            context.Database.ExecuteSqlRaw(sql);
        }
        StampVersion(context);
        transaction.Commit();
    }

    public static void StampVersion(AdressenDbContext context)
    {
        var sql = $"PRAGMA user_version = {AppSettings.DatabaseSchemaVersion};";
        context.Database.ExecuteSqlRaw(sql);
    }

    /// <summary>Konsistente Sicherungskopie über die SQLite-Backup-API (auch bei WAL-Modus korrekt) neben der Originaldatei; liefert den Pfad der Kopie.</summary>
    public static string CreateBackupCopy(string filePath, int fromVersion)
    {
        var directory = Path.GetDirectoryName(filePath) ?? string.Empty;
        var backupPath = Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(filePath)}.vor-Update-v{fromVersion}-{DateTime.Now:yyyyMMdd-HHmmss}{Path.GetExtension(filePath)}");
        using var source = new SqliteConnection($"Data Source={filePath}");
        using var target = new SqliteConnection($"Data Source={backupPath}");
        source.Open();
        target.Open();
        source.BackupDatabase(target);
        return backupPath;
    }

    // --- Hilfsmethoden ---

    /// <summary>Spalten der Tabelle Adressen aus dem EF-Modell (ohne Primärschlüssel) mit fertigem DDL-Fragment für ADD COLUMN.</summary>
    private static List<ColumnSpec> ModelColumns()
    {
        using var context = new AdressenDbContext(":memory:");  // nur für das Modell, es wird keine Verbindung geöffnet
        var model = context.GetService<IDesignTimeModel>().Model;  // das Laufzeitmodell (context.Model) enthält Collation/Defaults nicht
        var entity = model.FindEntityType(typeof(Adresse)) ?? throw new InvalidOperationException("Entität Adresse fehlt im EF-Modell.");
        var table = StoreObjectIdentifier.Table(entity.GetTableName() ?? "Adressen");
        var sample = new Adresse();  // liefert die Vorgabewerte neuer Datensätze (z. B. Reminder = true) – das EF-Modell kennt keine DB-Defaults
        var result = new List<ColumnSpec>();
        foreach (var property in entity.GetProperties())
        {
            if (property.IsPrimaryKey()) { continue; }
            var name = property.GetColumnName(table) ?? property.Name;
            var ddl = new StringBuilder($"\"{name}\" {property.GetColumnType()}");
            if (property.GetCollation() is { Length: > 0 } collation) { ddl.Append(" COLLATE ").Append(collation); }
            if (!property.IsNullable) { ddl.Append(" NOT NULL DEFAULT ").Append(DefaultLiteral(property, sample)); }  // ADD COLUMN NOT NULL verlangt einen Default
            result.Add(new ColumnSpec(name, ddl.ToString()));
        }
        return result;
    }

    /// <summary>SQL-Literal für den Vorgabewert einer NOT-NULL-Spalte, abgeleitet vom Wert der Eigenschaft in einem frisch erzeugten Datensatz.</summary>
    private static string DefaultLiteral(IProperty property, Adresse sample)
    {
        var value = property.PropertyInfo?.GetValue(sample);
        return value switch
        {
            bool b => b ? "1" : "0",
            string s => $"'{s.Replace("'", "''")}'",
            int or long or short or byte or double or float or decimal => Convert.ToString(value, CultureInfo.InvariantCulture)!,
            null when property.ClrType == typeof(string) => "''",
            _ => "0"
        };
    }

    private static object? Scalar(SqliteConnection connection, string sql)
    {
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        return command.ExecuteScalar();
    }

    private static HashSet<string> Names(SqliteConnection connection, string sql)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        using var reader = command.ExecuteReader();
        while (reader.Read()) { names.Add(reader.GetString(0)); }
        return names;
    }
}
