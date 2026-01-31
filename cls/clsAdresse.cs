using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Adressen.cls;

[Table("Gruppen")]
public class Gruppe
{
    [Key]
    public int Id
    {
        get; set;
    }
    [Required]
    public string Name { get; set; } = string.Empty;
    public virtual ICollection<Adresse> Adressen { get; set; } = [];
}

[Table("Dokumente")]
public class Dokument
{
    [Key]
    public int Id
    {
        get; set;
    }
    [Required]
    public string Dateipfad { get; set; } = string.Empty;
    public int AdressId
    {
        get; set;
    }
    [ForeignKey("AdressId")]
    public virtual Adresse Adresse { get; set; } = null!;
}

[Table("Adressen")]
public class Adresse : IContactEntity
{
    private string? _searchTextCache;
    // --- 1. Eigenschaften in der gewünschten Anzeige-Reihenfolge ---

    public string? Anrede
    {
        get; set;
    }
    public string? Praefix
    {
        get; set;
    }
    public string? Nachname
    {
        get; set;
    }
    public string? Vorname
    {
        get; set;
    }
    public string? Zwischenname
    {
        get; set;
    }
    public string? Nickname
    {
        get; set;
    }
    public string? Suffix
    {
        get; set;
    }
    public string? Unternehmen
    {
        get; set;
    }
    public string? Position
    {
        get; set;
    }
    public string? Strasse
    {
        get; set;
    }
    public string? PLZ
    {
        get; set;
    }
    public string? Ort
    {
        get; set;
    }
    public string? Postfach
    {
        get; set;
    }   
    public string? Land
    {
        get; set;
    }
    public string? Betreff
    {
        get; set;
    }
    public string? Grussformel
    {
        get; set;
    }
    public string? Schlussformel
    {
        get; set;
    }
    public DateOnly? Geburtstag
    {
        get; set;
    }
    public string? Mail1
    {
        get; set;
    }
    public string? Mail2
    {
        get; set;
    }
    public string? Telefon1
    {
        get; set;
    }
    public string? Telefon2
    {
        get; set;
    }
    public string? Mobil
    {
        get; set;
    }
    public string? Fax
    {
        get; set;
    }
    public string? Internet
    {
        get; set;
    }
    public string? Notizen
    {
        get; set;
    }

    // Id (UniqueId) soll an letzter Stelle angezeigt werden
    [Key]
    [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
    public int Id
    {
        get; set;
    }


    // --- 2. Ausgeblendete Navigation Properties ---

    [Browsable(false)]
    public virtual ICollection<Gruppe> Gruppen { get; set; } = [];

    [Browsable(false)]
    public virtual ICollection<Dokument> Dokumente { get; set; } = [];

    [Browsable(false)]
    public virtual Foto? Foto
    {
        get; set;
    }


    // --- 3. IContactEntity Implementierung (Ausgeblendet für Grid) ---

    // Wir blenden die Interface-Properties aus, da sie im Grid nur Duplikate wären
    // oder nicht angezeigt werden sollen.

    [NotMapped]
    [Browsable(false)]
    public string UniqueId => Id.ToString();

    [NotMapped]
    [Browsable(false)] // Soll nicht angezeigt werden
    public string DisplayName => $"{Vorname} {Nachname}".Trim();

    [NotMapped]
    [Browsable(false)]
    public string SearchText // Wenn Cache leer ist, berechnen (Lazy Loading)
    {
        get
        {
            _searchTextCache ??= $"{Vorname} {Nachname} {Unternehmen} {Position} {Ort} {PLZ} {Strasse} {Nickname} {Telefon1} {Telefon2} {Mobil} {Mail1} {Mail2} {Notizen} {Internet}".ToLowerInvariant();
            return _searchTextCache;
        }
    }

    [NotMapped]
    [Browsable(false)]
    public DateOnly? BirthdayDate => Geburtstag; //?.ToDateTime(TimeOnly.MinValue);

    [NotMapped]
    [Browsable(false)]
    public IList<string> GroupList => [.. Gruppen.Select(g => g.Name)];

    // --- 4. Methoden ---
    // Methode zum Zurücksetzen (wird beim Speichern aufgerufen)
    public void ResetSearchCache()
    {
        _searchTextCache = null;
    }

    public Task<Image?> GetPhotoAsync()
    {
        if (Foto?.Fotodaten == null) { return Task.FromResult<Image?>(null); }
        try
        {
            using var ms = new MemoryStream(Foto.Fotodaten);
            // Erstellt eine Kopie (Deep Copy), damit der MemoryStream geschlossen werden kann
            using var temp = Image.FromStream(ms);
            return Task.FromResult<Image?>(new Bitmap(temp));
        }
        catch { return Task.FromResult<Image?>(null); }
    }
}