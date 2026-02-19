using System.ComponentModel;
using System.Text;

namespace Adressen.cls;

[AttributeUsage(AttributeTargets.Property)]
internal class GoogleFieldAttribute(string category) : Attribute
{
    public string Category { get; } = category;
}

internal class Contact : ICloneable, IContactEntity
{
    private string? _searchTextCache;

    // ========================================================================
    // 1. EIGENSCHAFTEN MIT MAPPING-ATTRIBUTEN
    // ========================================================================

    [GoogleField("userDefined")]
    public string? Anrede
    {
        get; set;
    }
    [GoogleField("names")]
    public string? Praefix
    {
        get; set;
    }
    [GoogleField("names")]
    public string? Nachname
    {
        get; set;
    }
    [GoogleField("names")]
    public string? Vorname
    {
        get; set;
    }
    [GoogleField("names")]
    public string? Zwischenname
    {
        get; set;
    }
    [GoogleField("nicknames")]
    public string? Nickname
    {
        get; set;
    }
    [GoogleField("names")]
    public string? Suffix
    {
        get; set;
    }
    [GoogleField("organizations")]
    public string? Unternehmen
    {
        get; set;
    }
    [GoogleField("organizations")]
    public string? Position
    {
        get; set;
    }
    [GoogleField("addresses")]
    public string? Strasse
    {
        get; set;
    }
    [GoogleField("addresses")]
    public string? PLZ
    {
        get; set;
    }
    [GoogleField("addresses")]
    public string? Ort
    {
        get; set;
    }
    [GoogleField("addresses")]
    public string? Postfach
    {
        get; set;
    }
    [GoogleField("addresses")]
    public string? Land
    {
        get; set;
    }
    [GoogleField("userDefined")]
    public string? Betreff
    {
        get; set;
    }
    [GoogleField("userDefined")]
    public string? Grussformel
    {
        get; set;
    }
    [GoogleField("userDefined")]
    public string? Schlussformel
    {
        get; set;
    }
    [GoogleField("birthdays")]
    public DateOnly? Geburtstag
    {
        get; set;
    }
    [GoogleField("emailAddresses")]
    public string? Mail1
    {
        get; set;
    }
    [GoogleField("emailAddresses")]
    public string? Mail2
    {
        get; set;
    }
    [GoogleField("phoneNumbers")]
    public string? Telefon1
    {
        get; set;
    }
    [GoogleField("phoneNumbers")]
    public string? Telefon2
    {
        get; set;
    }
    [GoogleField("phoneNumbers")]
    public string? Mobil
    {
        get; set;
    }
    [GoogleField("phoneNumbers")]
    public string? Fax
    {
        get; set;
    }
    [GoogleField("urls")]
    public string? Internet
    {
        get; set;
    }
    [GoogleField("biographies")]
    public string? Notizen
    {
        get; set;
    }

    // Eigenschaften ohne Attribut (werden manuell oder gar nicht geprüft)
    public string ResourceName { get; set; } = string.Empty;

    // ========================================================================
    // 2. HILFS-PROPERTIES (Browsable false)
    // ========================================================================

    [Browsable(false)] public List<string> GroupNames { get; set; } = [];
    [Browsable(false)]
    public string? PhotoUrl
    {
        get; set;
    }
    [Browsable(false)] public string ETag { get; set; } = string.Empty;

    // ========================================================================
    // 3. IContactEntity IMPLEMENTIERUNG
    // ========================================================================

    [Browsable(false)] public string UniqueId => ResourceName;
    [Browsable(false)] public string DisplayName => $"{Vorname} {Nachname}".Trim();
    [Browsable(false)] public IList<string> GroupList => GroupNames;
    [Browsable(false)] public DateOnly? BirthdayDate => Geburtstag;

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
                sb.Append(Telefon1).Append(' ').Append(Telefon2).Append(' ').Append(Mobil).Append(' ');
                sb.Append(Mail1).Append(' ').Append(Mail2).Append(' ');
                sb.Append(Notizen).Append(' ').Append(Internet);
                _searchTextCache = sb.ToString().ToLowerInvariant();
            }
            return _searchTextCache;
        }
    }

    // ========================================================================
    // 4. METHODEN (Refactored & Vereinfacht)
    // ========================================================================

    public void ResetSearchCache() => _searchTextCache = null;

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
        clone.GroupNames = [.. GroupNames];
        return clone;
    }

    public void CopyFrom(Contact other)
    {
        if (other == null) { return; }

        // 1. Alle Standard-Properties kopieren
        var props = typeof(Contact).GetProperties();
        foreach (var prop in props)
        {
            // Wir kopieren nur, wenn man schreiben und lesen kann und es keine Liste ist
            if (prop.CanWrite && prop.CanRead && prop.Name != nameof(GroupNames)) { prop.SetValue(this, prop.GetValue(other)); }
        }

        // 2. Listen und Spezialfelder manuell behandeln (Deep Copy)
        GroupNames.Clear();
        if (other.GroupNames != null) { GroupNames.AddRange(other.GroupNames); }

        ResetSearchCache();
    }

    // --- AUTOMATISCH: Nutzt die [GoogleField] Attribute zur Erkennung ---
    public List<string> GetChangedFields(Contact original)
    {
        if (original == null) { return []; }

        var changes = new HashSet<string>(); // HashSet verhindert Duplikate automatisch
        var props = typeof(Contact).GetProperties();

        foreach (var prop in props)
        {
            // Hat die Property unser [GoogleField] Attribut?
            if (Attribute.GetCustomAttribute(prop, typeof(GoogleFieldAttribute)) is GoogleFieldAttribute attr)
            {
                var valCurrent = prop.GetValue(this);
                var valOriginal = prop.GetValue(original);
                if (!Equals(valCurrent, valOriginal)) { changes.Add(attr.Category); }
            }
        }

        // Spezialfälle prüfen (die kein einfaches Attribut haben)
        //if (PhotoUrl != original.PhotoUrl) { changes.Add("photos"); }
        if (!GroupNames.OrderBy(x => x).SequenceEqual(original.GroupNames.OrderBy(x => x))) { changes.Add("memberships"); }
        return [.. changes];
    }
}
