using System.ComponentModel;

namespace Adressen.cls;

internal class Contact : ICloneable, IContactEntity
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

    // ResourceName (UniqueId) an letzter Stelle
    public string ResourceName { get; set; } = string.Empty;


    // --- 2. Ausgeblendete Hilfs-Properties ---

    [Browsable(false)]
    public List<string> GroupNames { get; set; } = [];

    [Browsable(false)]
    public string? PhotoUrl
    {
        get; set;
    }

    [Browsable(false)]
    public string ETag { get; set; } = string.Empty;


    // --- 3. IContactEntity Implementierung (Ausgeblendet) ---

    [Browsable(false)]
    public string UniqueId => ResourceName;

    [Browsable(false)] // Soll nicht angezeigt werden
    public string DisplayName => $"{Vorname} {Nachname}".Trim();

    [Browsable(false)]
    public string SearchText
    {
        get
        {
            _searchTextCache ??= $"{Vorname} {Nachname} {Unternehmen} {Position} {Ort} {PLZ} {Strasse} {Nickname} {Telefon1} {Telefon2} {Mobil} {Mail1} {Mail2} {Notizen} {Internet}".ToLowerInvariant();
            return _searchTextCache;
        }
    }

    [Browsable(false)]
    public DateOnly? BirthdayDate => Geburtstag;

    [Browsable(false)]
    public IList<string> GroupList => GroupNames;

    // --- 4. Methoden ---
    public void ResetSearchCache()
    {
        _searchTextCache = null;
    }

    public async Task<Image?> GetPhotoAsync()
    {
        if (string.IsNullOrEmpty(PhotoUrl)) { return null; }

        try
        {
            var bytes = await HttpService.Client.GetByteArrayAsync(PhotoUrl);
            using var ms = new MemoryStream(bytes);
            return new Bitmap(ms);
        }
        catch { return null; }
    }

    public object Clone()
    {
        var clone = (Contact)MemberwiseClone();
        clone.GroupNames = [.. GroupNames]; // Erstelle eine neue Liste mit denselben Einträgen für den Klon
        return clone;
    }

    public List<string> GetChangedFields(Contact original)
    {
        var changes = new List<string>();

        static bool IsChanged(string? s1, string? s2)
        {
            return (s1 ?? string.Empty) != (s2 ?? string.Empty);
        }

        if (IsChanged(Vorname, original.Vorname) ||
            IsChanged(Nachname, original.Nachname) ||
            IsChanged(Praefix, original.Praefix) ||
            IsChanged(Zwischenname, original.Zwischenname) ||
            IsChanged(Suffix, original.Suffix))
        {
            changes.Add("names");
        }

        if (IsChanged(Nickname, original.Nickname)) { changes.Add("nicknames"); }
        if (IsChanged(Unternehmen, original.Unternehmen) || IsChanged(Position, original.Position)) { changes.Add("organizations"); }

        if (IsChanged(Strasse, original.Strasse) ||
            IsChanged(PLZ, original.PLZ) ||
            IsChanged(Ort, original.Ort) ||
            IsChanged(Postfach, original.Postfach) ||
            IsChanged(Land, original.Land))
        {
            changes.Add("addresses");
        }

        if (IsChanged(Mail1, original.Mail1) || IsChanged(Mail2, original.Mail2)) { changes.Add("emailAddresses"); }

        if (IsChanged(Telefon1, original.Telefon1) ||
            IsChanged(Telefon2, original.Telefon2) ||
            IsChanged(Mobil, original.Mobil) ||
            IsChanged(Fax, original.Fax))
        {
            changes.Add("phoneNumbers");
        }

        if (IsChanged(Notizen, original.Notizen)) { changes.Add("biographies"); }
        if (IsChanged(Internet, original.Internet)) { changes.Add("urls"); }
        if (Geburtstag != original.Geburtstag) { changes.Add("birthdays"); }

        if (IsChanged(Anrede, original.Anrede) ||
            IsChanged(Betreff, original.Betreff) ||
            IsChanged(Grussformel, original.Grussformel) ||
            IsChanged(Schlussformel, original.Schlussformel))
        {
            changes.Add("userDefined");
        }

        if (IsChanged(PhotoUrl, original.PhotoUrl)) { changes.Add("photos"); }

        if (!GroupNames.OrderBy(static x => x).SequenceEqual(original.GroupNames.OrderBy(static x => x)))
        {
            changes.Add("memberships");
        }

        return [.. changes.Distinct()];
    }
}