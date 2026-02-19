using Microsoft.EntityFrameworkCore;

namespace Adressen.cls;

internal class AdressenDbContext(string dbPath) : DbContext
{
    public DbSet<Adresse> Adressen { get; set; } = null!;
    // Hinweis: Die DbSet<Foto> ist eigentlich optional, da sie über Adresse.Foto erreichbar ist, 
    // aber es schadet nicht, sie direkt abrufbar zu haben.
    public DbSet<Foto> Fotos { get; set; } = null!;
    public DbSet<Gruppe> Gruppen { get; set; } = null!;
    public DbSet<Dokument> Dokumente { get; set; } = null!;

    private readonly string _dbPath = dbPath;

    protected override void OnConfiguring(DbContextOptionsBuilder options)
        => options.UseSqlite($"Data Source={_dbPath}");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        // 1. Globales Verhalten: Alle Strings case-insensitive (NOCASE) für SQLite
        foreach (var entityType in modelBuilder.Model.GetEntityTypes())
        {
            foreach (var property in entityType.GetProperties())
            {
                if (property.ClrType == typeof(string))
                {
                    property.SetCollation("NOCASE");
                }
            }
        }

        // 2. Beziehungen konfigurieren

        // 1:1 Foto Beziehung
        modelBuilder.Entity<Adresse>()
            .HasOne(a => a.Foto)
            .WithOne(f => f.Adresse)
            .HasForeignKey<Foto>(f => f.AdressId)
            .OnDelete(DeleteBehavior.Cascade); // Foto wird gelöscht, wenn Adresse gelöscht wird

        // 1:N Dokumente Beziehung
        modelBuilder.Entity<Adresse>()
            .HasMany(a => a.Dokumente)
            .WithOne(d => d.Adresse)
            .HasForeignKey(d => d.AdressId)
            .OnDelete(DeleteBehavior.Cascade); // Dokumente werden gelöscht, wenn Adresse gelöscht wird

        // M:N Gruppen Beziehung
        modelBuilder.Entity<Adresse>()
            .HasMany(a => a.Gruppen)
            .WithMany(g => g.Adressen)
            .UsingEntity(
                "AdresseGruppen", // Expliziter Name der Verknüpfungstabelle
                                  // Konfiguration der FKs zur Gruppe (Right)
                l => l.HasOne(typeof(Gruppe))
                      .WithMany()
                      .HasForeignKey("GruppenId")
                      .HasPrincipalKey(nameof(Gruppe.Id)),
                // Konfiguration der FKs zur Adresse (Left)
                r => r.HasOne(typeof(Adresse))
                      .WithMany()
                      .HasForeignKey("AdressenId")
                      .HasPrincipalKey(nameof(Adresse.Id)),
                // Konfiguration des PKs der Verknüpfungstabelle
                j => j.HasKey("AdressenId", "GruppenId")
            );
    }
}
