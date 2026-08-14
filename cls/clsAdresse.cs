using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

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
    [MaxLength(100)] // Ein sinnvolles Limit für Gruppennamen
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
    [MaxLength(1000)] // Windows-Pfade können lang sein, 1000 ist hier sehr sicher
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

    [MaxLength(50)]
    public string? Anrede
    {
        get; set;
    }

    [MaxLength(50)]
    [DisplayName("Titel")]
    public string? Praefix
    {
        get; set;
    }

    [MaxLength(100)]
    public string? Nachname
    {
        get; set;
    }

    [MaxLength(100)]
    public string? Vorname
    {
        get; set;
    }

    [MaxLength(100)]
    public string? Zwischenname
    {
        get; set;
    }

    [MaxLength(50)]
    public string? Nickname
    {
        get; set;
    }

    [MaxLength(50)]
    public string? Suffix
    {
        get; set;
    }

    [MaxLength(150)]
    public string? Unternehmen
    {
        get; set;
    }

    [MaxLength(100)]
    public string? Position
    {
        get; set;
    }

    [MaxLength(150)]
    [DisplayName("Adresse")]
    public string? Strasse
    {
        get; set;
    }

    [MaxLength(20)]
    public string? PLZ
    {
        get; set;
    }

    [MaxLength(100)]
    public string? Ort
    {
        get; set;
    }

    [MaxLength(50)]
    public string? Postfach
    {
        get; set;
    }

    [MaxLength(100)]
    public string? Land
    {
        get; set;
    }

    [MaxLength(150)]
    public string? Betreff
    {
        get; set;
    }

    [MaxLength(100)]
    [DisplayName("Grußformel")]
    public string? Grussformel
    {
        get; set;
    }

    [MaxLength(100)]
    public string? Schlussformel
    {
        get; set;
    }

    public DateOnly? Geburtstag
    {
        get; set;
    }

    [MaxLength(254)]
    public string? Mail1
    {
        get; set;
    }

    [MaxLength(254)]
    public string? Mail2
    {
        get; set;
    }

    [MaxLength(50)]
    public string? Telefon1
    {
        get; set;
    }

    [MaxLength(50)]
    public string? Telefon2
    {
        get; set;
    }

    [MaxLength(50)]
    public string? Mobil
    {
        get; set;
    }

    [MaxLength(50)]
    public string? Fax
    {
        get; set;
    }

    [MaxLength(2048)]
    public string? Internet
    {
        get; set;
    }

    [MaxLength(1000)] // Analog zur Contact-Klasse
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

    //[Browsable(false)]  // auskommentiert lassen — bricht DataBinding; stattdessen: if (dgvAdressen.Columns["Reminder"] is DataGridViewColumn col) col.Visible = false;
    public bool Reminder { get; set; } = true; // Standardmäßig aktivieren wir die Erinnerung für bestehende/neue Adressen

    [Browsable(false)]
    public virtual ICollection<Gruppe> Gruppen { get; set; } = [];

    [Browsable(false)]
    public virtual ICollection<Dokument> Dokumente { get; set; } = [];

    [Browsable(false)]
    public virtual Foto? Foto
    {
        get; set;
    }

    [Browsable(false)]
    public DateTime? LastModified { get; set; } = DateTime.UtcNow;

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
    public string SearchText
    {
        get
        {
            if (_searchTextCache == null)
            {
                var sb = new StringBuilder();
                sb.Append(Vorname).Append(' ').Append(Nachname).Append(' ');
                sb.Append(Unternehmen).Append(' ').Append(Position).Append(' ');
                sb.Append(Ort).Append(' ').Append(PLZ).Append(' ').Append(Strasse).Append(' ');
                sb.Append(Nickname).Append(' ');
                // Telefonnummern: roh (für Teilsuche) + normalisiert (für Fritz-Monitor-Treffer)
                foreach (var num in new[] { Telefon1, Telefon2, Mobil, Fax })
                {
                    if (!string.IsNullOrWhiteSpace(num))
                    {
                        sb.Append(num).Append(' ');
                        sb.Append(Utils.NormalizePhoneNumber(num)).Append(' ');
                    }
                }
                sb.Append(Mail1).Append(' ').Append(Mail2).Append(' ');
                sb.Append(Notizen).Append(' ').Append(Internet);
                _searchTextCache = sb.ToString().ToLowerInvariant();
            }
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

    public Task<Image?> GetPhotoAsync(CancellationToken token = default)
    {
        if (token.IsCancellationRequested) { return Task.FromResult<Image?>(null); }  // Sofortiger Abbruch, falls der Nutzer bereits weitergescrollt hat
        if (Foto?.Fotodaten == null) { return Task.FromResult<Image?>(null); }
        try
        {
            using var ms = new MemoryStream(Foto.Fotodaten);
            using var temp = Image.FromStream(ms);
            return Task.FromResult<Image?>(new Bitmap(temp));
        }
        catch { return Task.FromResult<Image?>(null); }
    }

    public object? GetPropertyValue(string propertyName)
    {
        return propertyName switch
        {
            nameof(Id) => Id,
            nameof(Anrede) => Anrede,
            nameof(Praefix) => Praefix,
            nameof(Nachname) => Nachname,
            nameof(Vorname) => Vorname,
            nameof(Zwischenname) => Zwischenname,
            nameof(Nickname) => Nickname,
            nameof(Suffix) => Suffix,
            nameof(Unternehmen) => Unternehmen,
            nameof(Position) => Position,
            nameof(Strasse) => Strasse,
            nameof(PLZ) => PLZ,
            nameof(Ort) => Ort,
            nameof(Postfach) => Postfach,
            nameof(Land) => Land,
            nameof(Betreff) => Betreff,
            nameof(Grussformel) => Grussformel,
            nameof(Schlussformel) => Schlussformel,
            nameof(Geburtstag) => Geburtstag,
            nameof(Reminder) => Reminder,
            nameof(Mail1) => Mail1,
            nameof(Mail2) => Mail2,
            nameof(Telefon1) => Telefon1,
            nameof(Telefon2) => Telefon2,
            nameof(Mobil) => Mobil,
            nameof(Fax) => Fax,
            nameof(Internet) => Internet,
            nameof(Notizen) => Notizen,
            _ => null
        };
    }

    public void SetPropertyValue(string propertyName, string? value)
    {
        switch (propertyName)
        {
            case nameof(Anrede): Anrede = value; break;
            case nameof(Praefix): Praefix = value; break;
            case nameof(Nachname): Nachname = value; break;
            case nameof(Vorname): Vorname = value; break;
            case nameof(Zwischenname): Zwischenname = value; break;
            case nameof(Nickname): Nickname = value; break;
            case nameof(Suffix): Suffix = value; break;
            case nameof(Unternehmen): Unternehmen = value; break;
            case nameof(Position): Position = value; break;
            case nameof(Strasse): Strasse = value; break;
            case nameof(PLZ): PLZ = value; break;
            case nameof(Ort): Ort = value; break;
            case nameof(Postfach): Postfach = value; break;
            case nameof(Land): Land = value; break;
            case nameof(Betreff): Betreff = value; break;
            case nameof(Grussformel): Grussformel = value; break;
            case nameof(Schlussformel): Schlussformel = value; break;
            case nameof(Mail1): Mail1 = value; break;
            case nameof(Mail2): Mail2 = value; break;
            case nameof(Telefon1): Telefon1 = value; break;
            case nameof(Telefon2): Telefon2 = value; break;
            case nameof(Mobil): Mobil = value; break;
            case nameof(Fax): Fax = value; break;
            case nameof(Internet): Internet = value; break;
            case nameof(Notizen): Notizen = value; break;
            case nameof(Reminder):
                if (bool.TryParse(value, out var bValue)) { Reminder = bValue; }
                else if (value == "1") { Reminder = true; }
                else if (value == "0") { Reminder = false; }
                break;
            default: break; // Unbekannte Spalten einfach ignorieren
        }
    }

    public void TrimStrings()  // Trimming vor dem Speichern
    {
        Anrede = Anrede?.Trim();
        Praefix = Praefix?.Trim();
        Nachname = Nachname?.Trim();
        Vorname = Vorname?.Trim();
        Zwischenname = Zwischenname?.Trim();
        Nickname = Nickname?.Trim();
        Suffix = Suffix?.Trim();
        Unternehmen = Unternehmen?.Trim();
        Position = Position?.Trim();
        Strasse = Strasse?.Trim();
        PLZ = PLZ?.Trim();
        Ort = Ort?.Trim();
        Postfach = Postfach?.Trim();
        Land = Land?.Trim();
        Betreff = Betreff?.Trim();
        Grussformel = Grussformel?.Trim();
        Schlussformel = Schlussformel?.Trim();
        Mail1 = Mail1?.Trim();
        Mail2 = Mail2?.Trim();
        Telefon1 = Telefon1?.Trim();
        Telefon2 = Telefon2?.Trim();
        Mobil = Mobil?.Trim();
        Fax = Fax?.Trim();
        Internet = Internet?.Trim();
        Notizen = Notizen?.Trim();
    }
}