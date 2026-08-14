using Microsoft.EntityFrameworkCore;

namespace Adressen.cls;

internal class AdressenDbContext(string dbPath) : DbContext
{
    public DbSet<Adresse> Adressen { get; set; } = null!;
    public DbSet<Gruppe> Gruppen { get; set; } = null!;
    public DbSet<Dokument> Dokumente { get; set; } = null!;

    protected override void OnConfiguring(DbContextOptionsBuilder options) => options.UseSqlite($"Data Source={dbPath}");
    protected override void ConfigureConventions(ModelConfigurationBuilder configurationBuilder) => configurationBuilder.Properties<string>().UseCollation("NOCASE"); // Setzt 'NOCASE' global für alle string-Eigenschaften im gesamten Modell

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
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

    public async override Task<int> SaveChangesAsync(CancellationToken cancellationToken = default)
    {
        SetLastModified();
        return await base.SaveChangesAsync(cancellationToken);
    }


    public override int SaveChanges()  // Synchrone Variante absichern (wird z.B. im Migrator verwendet)
    {
        SetLastModified();
        return base.SaveChanges();
    }

    private void SetLastModified()
    {
        // Direkt geänderte Adressen
        var changedAddresses = ChangeTracker.Entries<Adresse>().Where(e => e.State is EntityState.Modified or EntityState.Added).Select(e => e.Entity).ToHashSet();

        // Indirekte Änderungen über Kind-Objekte: Fotos
        foreach (var entry in ChangeTracker.Entries<Foto>().Where(e => e.State is EntityState.Modified or EntityState.Added or EntityState.Deleted))
        {
            var adressId = entry.State is EntityState.Modified or EntityState.Deleted ? (int)entry.OriginalValues[nameof(Foto.AdressId)]! : entry.Entity.AdressId;
            var addr = Set<Adresse>().Local.FirstOrDefault(a => a.Id == adressId);
            if (addr != null) { changedAddresses.Add(addr); }
        }

        // Indirekte Änderungen über Kind-Objekte: Dokumente
        foreach (var entry in ChangeTracker.Entries<Dokument>().Where(e => e.State is EntityState.Modified or EntityState.Added or EntityState.Deleted))
        {
            var adressId = entry.State == EntityState.Deleted ? (int)entry.OriginalValues[nameof(Dokument.AdressId)]! : entry.Entity.AdressId;
            var addr = Set<Adresse>().Local.FirstOrDefault(a => a.Id == adressId);
            if (addr != null) { changedAddresses.Add(addr); }
        }

        // Indirekte Änderungen über M:N Verknüpfungstabelle: Gruppen
        foreach (var entry in ChangeTracker.Entries().Where(e => e.Metadata.Name == "AdresseGruppen" && e.State is EntityState.Added or EntityState.Deleted))
        {
            // Wir greifen auf den in OnModelCreating definierten FK "AdressenId" zu
            var adressIdProp = entry.Properties.FirstOrDefault(p => p.Metadata.Name == "AdressenId");
            if (adressIdProp != null)
            {
                var adressId = entry.State == EntityState.Deleted ? (int)adressIdProp.OriginalValue! : (int)adressIdProp.CurrentValue!;
                var addr = Set<Adresse>().Local.FirstOrDefault(a => a.Id == adressId);  // Local: Funktioniert nur, wenn die Adresse bereits im aktuellen DbContext-Scope geladen wurde                                                                                        
                if (addr != null) { changedAddresses.Add(addr); }
            }
        }
        foreach (var adresse in changedAddresses)
        {
            adresse.LastModified = DateTime.UtcNow;
            adresse.TrimStrings();
        }
    }
}
