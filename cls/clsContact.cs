using System.Collections.Concurrent;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text;
using Google.Apis.PeopleService.v1.Data;

namespace Adressen.cls;

[AttributeUsage(AttributeTargets.Property)]
internal class GoogleFieldAttribute(string category) : Attribute
{
    public string Category { get; } = category;
}

internal class Contact : ICloneable, IContactEntity
{
    private string? _searchTextCache;
    private static readonly ConcurrentDictionary<string, byte[]> _photoCache = new(); // Statischer Cache, der für die gesamte Laufzeit der App existiert

    // ========================================================================
    // 1. EIGENSCHAFTEN MIT MAPPING-ATTRIBUTEN
    // ========================================================================

    [GoogleField("userDefined")]
    [MaxLength(50)]
    public string? Anrede
    {
        get; set;
    }

    [GoogleField("names")]
    [MaxLength(50)]
    [DisplayName("Präfix")]
    public string? Praefix
    {
        get; set;
    }

    [GoogleField("names")]
    [MaxLength(100)]
    public string? Nachname
    {
        get; set;
    }

    [GoogleField("names")]
    [MaxLength(100)]
    public string? Vorname
    {
        get; set;
    }

    [GoogleField("names")]
    [MaxLength(100)]
    public string? Zwischenname
    {
        get; set;
    }

    [GoogleField("nicknames")]
    [MaxLength(50)]
    public string? Nickname
    {
        get; set;
    }

    [GoogleField("names")]
    [MaxLength(50)]
    public string? Suffix
    {
        get; set;
    }

    [GoogleField("organizations")]
    [MaxLength(150)]
    public string? Unternehmen
    {
        get; set;
    }

    [GoogleField("organizations")]
    [MaxLength(100)]
    public string? Position
    {
        get; set;
    }

    [GoogleField("addresses")]
    [MaxLength(150)]
    [DisplayName("Straße")]
    public string? Strasse
    {
        get; set;
    }

    [GoogleField("addresses")]
    [MaxLength(20)]
    public string? PLZ
    {
        get; set;
    }

    [GoogleField("addresses")]
    [MaxLength(100)]
    public string? Ort
    {
        get; set;
    }

    [GoogleField("addresses")]
    [MaxLength(50)]
    public string? Postfach
    {
        get; set;
    }

    [GoogleField("addresses")]
    [MaxLength(100)]
    public string? Land
    {
        get; set;
    }

    [GoogleField("userDefined")]
    [MaxLength(150)]
    public string? Betreff
    {
        get; set;
    }

    [GoogleField("userDefined")]
    [MaxLength(100)]
    [DisplayName("Grußformel")]
    public string? Grussformel
    {
        get; set;
    }

    [GoogleField("userDefined")]
    [MaxLength(100)]
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
    [MaxLength(254)] // Strenger RFC-Standard für E-Mails
    public string? Mail1
    {
        get; set;
    }

    [GoogleField("emailAddresses")]
    [MaxLength(254)]
    public string? Mail2
    {
        get; set;
    }

    [GoogleField("phoneNumbers")]
    [MaxLength(50)]
    public string? Telefon1
    {
        get; set;
    }

    [GoogleField("phoneNumbers")]
    [MaxLength(50)]
    public string? Telefon2
    {
        get; set;
    }

    [GoogleField("phoneNumbers")]
    [MaxLength(50)]
    public string? Mobil
    {
        get; set;
    }

    [GoogleField("phoneNumbers")]
    [MaxLength(50)]
    public string? Fax
    {
        get; set;
    }

    [GoogleField("urls")]
    [MaxLength(2048)] // Ausreichend für extrem lange Links
    public string? Internet
    {
        get; set;
    }

    [GoogleField("biographies")]
    [MaxLength(1000)] // Das besprochene sichere Sync-Limit für Smartphones
    public string? Notizen
    {
        get; set;
    }

    // Eigenschaften ohne Attribut (werden manuell oder gar nicht geprüft)
    [MaxLength(200)] // Auch hier eine Obergrenze für die DB sinnvoll
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

    [Browsable(false)]
    [Newtonsoft.Json.JsonIgnore]
    public Person? RawGooglePerson
    {
        get; set;
    }

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

    //public async Task<Image?> GetPhotoAsync()
    //{
    //    if (string.IsNullOrEmpty(PhotoUrl)) { return null; }

    //    try
    //    {
    //        // 1. Prüfen, ob das Bild schon im Speicher liegt
    //        if (!_photoCache.TryGetValue(PhotoUrl, out var bytes))
    //        {
    //            // 2. Falls nicht: Herunterladen und sicher im Cache ablegen
    //            bytes = await HttpService.Client.GetByteArrayAsync(PhotoUrl);
    //            _photoCache.TryAdd(PhotoUrl, bytes);
    //        }

    //        using var ms = new MemoryStream(bytes);
    //        return new Bitmap(ms);
    //    }
    //    catch { return null; }
    //}

    public async Task<Image?> GetPhotoAsync(CancellationToken token = default)
    {
        if (string.IsNullOrEmpty(PhotoUrl))
        {
            return null;
        }

        try
        {
            // 1. Prüfen, ob das Bild schon im Speicher liegt
            if (!_photoCache.TryGetValue(PhotoUrl, out var bytes))
            {
                // 2. Falls nicht: Herunterladen und Token durchreichen!
                // Der HttpClient bricht den Web-Request bei einem Scroll-Event sofort ab.
                bytes = await HttpService.Client.GetByteArrayAsync(PhotoUrl, token);
                _photoCache.TryAdd(PhotoUrl, bytes);
            }

            // 3. Wenn während des Wartens auf den Cache oder den Stream abgebrochen wurde:
            if (token.IsCancellationRequested) { return null; }

            using var ms = new MemoryStream(bytes);
            using var temp = Image.FromStream(ms);
            return new Bitmap(temp);  // Deep Copy, damit der MemoryStream sofort geschlossen werden kann
        }
        catch (TaskCanceledException)
        {
            // Der Download wurde durch das schnelle Scrollen im DataGridView planmäßig abgebrochen
            return null;
        }
        catch
        {
            return null;
        }
    }

    public object Clone()
    {
        var clone = (Contact)MemberwiseClone();
        clone.GroupNames = [.. GroupNames];
        if (RawGooglePerson != null)  // Deep Clone für das Google-Objekt, damit der Snapshot unabhängig bleibt
        {
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(RawGooglePerson);
            clone.RawGooglePerson = Newtonsoft.Json.JsonConvert.DeserializeObject<Person>(json);
        }
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
