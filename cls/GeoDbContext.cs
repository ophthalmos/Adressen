using Microsoft.EntityFrameworkCore;

namespace Adressen.cls;


public class GeoDbContext : DbContext
{
    public static readonly string DbPath = Path.Combine(Path.GetDirectoryName(Environment.ProcessPath) ?? AppDomain.CurrentDomain.BaseDirectory, "geodata.db");

    public static bool DatabaseExists => File.Exists(DbPath);

    public DbSet<GeoStrasse> GeoStrassen { get; set; } = null!;

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        var uri = new Uri(DbPath).AbsoluteUri;                                  // kümmert sich korrekt um Sonderzeichen, Leerzeichen etc.
        _ = optionsBuilder.UseSqlite($"Data Source={uri}?mode=ro&immutable=1")  // immutable=1 für "Program Files"; keine Wal-Datei, keine Sperren, nur lesend
         .UseQueryTrackingBehavior(QueryTrackingBehavior.NoTracking);       // erspart AsNoTracking() in der StrassenAbfrage (query = geoContext.GeoStrassen…)
    }
}

public class GeoStrasse
{
    public int Id
    {
        get; set;
    }
    public string Strasse { get; set; } = string.Empty;
    public string PLZ { get; set; } = string.Empty;
    public string Ort { get; set; } = string.Empty;
    public string Bundesland { get; set; } = string.Empty;
}