using Microsoft.EntityFrameworkCore;

namespace Adressen.cls;

internal class AdressenDbContext(string dbPath) : DbContext
{
    public DbSet<Adresse> Adressen { get; set; } = null!;
    public DbSet<Foto> Fotos { get; set; } = null!;
    public DbSet<Gruppe> Gruppen { get; set; } = null!;
    public DbSet<Dokument> Dokumente { get; set; } = null!;

    private readonly string _dbPath = dbPath;

    protected override void OnConfiguring(DbContextOptionsBuilder options) => options.UseSqlite($"Data Source={_dbPath}");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        // Case-Insensitivity für SQLite konfigurieren
        modelBuilder.Entity<Adresse>(entity =>
        {
            entity.Property(e => e.Nachname).UseCollation("NOCASE");
            entity.Property(e => e.Vorname).UseCollation("NOCASE");
            entity.Property(e => e.Unternehmen).UseCollation("NOCASE");
            entity.Property(e => e.Ort).UseCollation("NOCASE");
            entity.Property(e => e.Strasse).UseCollation("NOCASE");
        });
        // 1:1 Foto Beziehung
        modelBuilder.Entity<Adresse>()
            .HasOne(a => a.Foto)
            .WithOne(f => f.Adresse)
            .HasForeignKey<Foto>(f => f.AdressId)
            .OnDelete(DeleteBehavior.Cascade);

        // 1:N Dokumente Beziehung
        modelBuilder.Entity<Adresse>()
            .HasMany(a => a.Dokumente)
            .WithOne(d => d.Adresse)
            .HasForeignKey(d => d.AdressId)
            .OnDelete(DeleteBehavior.Cascade); // Löscht Doku-Einträge, wenn Adresse gelöscht wird

        // M:N Gruppen Beziehung
        modelBuilder.Entity<Adresse>()
            .HasMany(a => a.Gruppen)
            .WithMany(g => g.Adressen)
            .UsingEntity(
                "AdresseGruppen", // Name der Tabelle
                l => l.HasOne(typeof(Gruppe)).WithMany().HasForeignKey("GruppenId").HasPrincipalKey(nameof(Gruppe.Id)),
                r => r.HasOne(typeof(Adresse)).WithMany().HasForeignKey("AdressenId").HasPrincipalKey(nameof(Adresse.Id)),
                j => j.HasKey("AdressenId", "GruppenId") // Primärschlüssel definieren
            );
    }
}
