using System.Globalization;
using System.Text;

namespace Adressen.cls;

internal static class VCardService
{

    internal sealed class VCardImportResult  // Ergebnis eines vCard-Imports: die befüllte Entität plus Rohdaten,  die erst im Form-Kontext aufgelöst werden können (Gruppen-Lookup, Foto-Bytes).
    {
        public required Adresse? Adresse
        {
            get; init;
        }
        public required Contact? Contact
        {
            get; init;
        }
        public List<string> GruppenNamen { get; init; } = [];
        public byte[]? FotoBytes      // nur für Adresse relevant
        {
            get; init;
        }
    }

    public static string ExportAdresse(Adresse a, bool includeFoto = false)
    {
        var sb = new StringBuilder();
        Begin(sb);
        sb.AppendLine($"N:{Esc(a.Nachname)};{Esc(a.Vorname)};{Esc(a.Zwischenname)};{Esc(a.Praefix)};{Esc(a.Suffix)}");  // Name
        sb.AppendLine($"FN:{Esc(BuildFN(a.Vorname, a.Zwischenname, a.Nachname, a.Unternehmen))}");
        if (!string.IsNullOrWhiteSpace(a.Nickname)) { sb.AppendLine($"NICKNAME:{Esc(a.Nickname)}"); }
        if (!string.IsNullOrWhiteSpace(a.Anrede)) { sb.AppendLine($"X-ANREDE:{Esc(a.Anrede)}"); }
        if (!string.IsNullOrWhiteSpace(a.Unternehmen)) { sb.AppendLine($"ORG:{Esc(a.Unternehmen)}"); }
        if (!string.IsNullOrWhiteSpace(a.Position)) { sb.AppendLine($"TITLE:{Esc(a.Position)}"); }
        if (HasAny(a.Strasse, a.PLZ, a.Ort, a.Land, a.Postfach)) { sb.AppendLine($"ADR;TYPE=WORK:{Esc(a.Postfach)};;{Esc(a.Strasse)};{Esc(a.Ort)};;{Esc(a.PLZ)};{Esc(a.Land)}"); }
        if (!string.IsNullOrWhiteSpace(a.Telefon1)) { sb.AppendLine($"TEL;TYPE=WORK,VOICE:{Esc(a.Telefon1)}"); }
        if (!string.IsNullOrWhiteSpace(a.Telefon2)) { sb.AppendLine($"TEL;TYPE=WORK,VOICE:{Esc(a.Telefon2)}"); }
        if (!string.IsNullOrWhiteSpace(a.Mobil)) { sb.AppendLine($"TEL;TYPE=CELL:{Esc(a.Mobil)}"); }
        if (!string.IsNullOrWhiteSpace(a.Fax)) { sb.AppendLine($"TEL;TYPE=FAX:{Esc(a.Fax)}"); }
        if (!string.IsNullOrWhiteSpace(a.Mail1)) { sb.AppendLine($"EMAIL;TYPE=INTERNET:{Esc(a.Mail1)}"); }
        if (!string.IsNullOrWhiteSpace(a.Mail2)) { sb.AppendLine($"EMAIL;TYPE=INTERNET:{Esc(a.Mail2)}"); }
        if (!string.IsNullOrWhiteSpace(a.Internet)) { sb.AppendLine($"URL:{Esc(a.Internet)}"); }
        if (a.Geburtstag.HasValue) { sb.AppendLine($"BDAY:{a.Geburtstag.Value:yyyy-MM-dd}"); }
        if (!string.IsNullOrWhiteSpace(a.Notizen)) { sb.AppendLine($"NOTE:{EscNote(a.Notizen)}"); }
        var cats = string.Join(",", a.Gruppen.Select(g => g.Name).Where(n => !string.IsNullOrWhiteSpace(n)).Select(Esc));  // Gruppen → CATEGORIES
        if (!string.IsNullOrWhiteSpace(cats)) { sb.AppendLine($"CATEGORIES:{cats}"); }
        if (!string.IsNullOrWhiteSpace(a.Betreff)) { sb.AppendLine($"X-BETREFF:{Esc(a.Betreff)}"); }
        if (!string.IsNullOrWhiteSpace(a.Grussformel)) { sb.AppendLine($"X-GRUSSFORMEL:{Esc(a.Grussformel)}"); }
        if (!string.IsNullOrWhiteSpace(a.Schlussformel)) { sb.AppendLine($"X-SCHLUSSFORMEL:{Esc(a.Schlussformel)}"); }
        if (includeFoto && a.Foto?.Fotodaten is { Length: > 0 } fotoBytes)
        {
            var b64 = Convert.ToBase64String(fotoBytes);
            var chunks = Chunks(b64, 72).ToList();
            var imageType = DetectImageType(fotoBytes);
            sb.Append($"PHOTO;ENCODING=b;TYPE={imageType}:");
            sb.AppendLine(chunks[0]);
            foreach (var chunk in chunks.Skip(1)) { sb.AppendLine(" " + chunk); }
        }
        End(sb);
        return sb.ToString();
    }

    public static string ExportContact(Contact c, bool includePhotoUrl = false)
    {
        var sb = new StringBuilder();
        Begin(sb);
        sb.AppendLine($"N:{Esc(c.Nachname)};{Esc(c.Vorname)};{Esc(c.Zwischenname)};{Esc(c.Praefix)};{Esc(c.Suffix)}");
        sb.AppendLine($"FN:{Esc(BuildFN(c.Vorname, c.Zwischenname, c.Nachname, c.Unternehmen))}");
        if (!string.IsNullOrWhiteSpace(c.Nickname)) { sb.AppendLine($"NICKNAME:{Esc(c.Nickname)}"); }
        if (!string.IsNullOrWhiteSpace(c.Anrede)) { sb.AppendLine($"X-ANREDE:{Esc(c.Anrede)}"); }
        if (!string.IsNullOrWhiteSpace(c.Unternehmen)) { sb.AppendLine($"ORG:{Esc(c.Unternehmen)}"); }
        if (!string.IsNullOrWhiteSpace(c.Position)) { sb.AppendLine($"TITLE:{Esc(c.Position)}"); }
        if (HasAny(c.Strasse, c.PLZ, c.Ort, c.Land, c.Postfach)) { sb.AppendLine($"ADR;TYPE=WORK:{Esc(c.Postfach)};;{Esc(c.Strasse)};{Esc(c.Ort)};;{Esc(c.PLZ)};{Esc(c.Land)}"); }
        if (!string.IsNullOrWhiteSpace(c.Telefon1)) { sb.AppendLine($"TEL;TYPE=WORK,VOICE:{Esc(c.Telefon1)}"); }
        if (!string.IsNullOrWhiteSpace(c.Telefon2)) { sb.AppendLine($"TEL;TYPE=WORK,VOICE:{Esc(c.Telefon2)}"); }
        if (!string.IsNullOrWhiteSpace(c.Mobil)) { sb.AppendLine($"TEL;TYPE=CELL:{Esc(c.Mobil)}"); }
        if (!string.IsNullOrWhiteSpace(c.Fax)) { sb.AppendLine($"TEL;TYPE=FAX:{Esc(c.Fax)}"); }
        if (!string.IsNullOrWhiteSpace(c.Mail1)) { sb.AppendLine($"EMAIL;TYPE=INTERNET:{Esc(c.Mail1)}"); }
        if (!string.IsNullOrWhiteSpace(c.Mail2)) { sb.AppendLine($"EMAIL;TYPE=INTERNET:{Esc(c.Mail2)}"); }
        if (!string.IsNullOrWhiteSpace(c.Internet)) { sb.AppendLine($"URL:{Esc(c.Internet)}"); }
        if (c.Geburtstag.HasValue) { sb.AppendLine($"BDAY:{c.Geburtstag.Value:yyyy-MM-dd}"); }
        if (!string.IsNullOrWhiteSpace(c.Notizen)) { sb.AppendLine($"NOTE:{EscNote(c.Notizen)}"); }
        var cats = string.Join(",", c.GroupNames.Where(n => !string.IsNullOrWhiteSpace(n)).Select(Esc));
        if (!string.IsNullOrWhiteSpace(cats)) { sb.AppendLine($"CATEGORIES:{cats}"); }
        if (!string.IsNullOrWhiteSpace(c.Betreff)) { sb.AppendLine($"X-BETREFF:{Esc(c.Betreff)}"); }
        if (!string.IsNullOrWhiteSpace(c.Grussformel)) { sb.AppendLine($"X-GRUSSFORMEL:{Esc(c.Grussformel)}"); }
        if (!string.IsNullOrWhiteSpace(c.Schlussformel)) { sb.AppendLine($"X-SCHLUSSFORMEL:{Esc(c.Schlussformel)}"); }
        if (includePhotoUrl && !string.IsNullOrWhiteSpace(c.PhotoUrl)) { sb.AppendLine($"PHOTO;VALUE=URI:{c.PhotoUrl}"); }  // Google: Foto ist nur eine URL, kein Base64-Blob
        End(sb);
        return sb.ToString();
    }

    public static VCardImportResult Import(string vcfPath, bool asContact)
    {
        var lines = UnfoldLines(File.ReadAllLines(vcfPath, Encoding.UTF8)).ToList();  // vCard 2.1 (oft ISO-8859-1) ist UTF8 riskant, wir reparieren über Quoted-Printable-Dekodierung
        var gruppenNamen = new List<string>();
        byte[]? fotoBytes = null;
        var phoneCount = 0;
        var mailCount = 0;
        string? anrede = null, praefix = null, nachname = null, vorname = null,
                zwischen = null, nickname = null, suffix = null,
                unternehmen = null, position = null,
                strasse = null, plz = null, ort = null, postfach = null, land = null,
                tel1 = null, tel2 = null, mobil = null, fax = null,
                mail1 = null, mail2 = null,
                internet = null, notizen = null,
                betreff = null, gruss = null, schluss = null,
                photoUrl = null;
        DateOnly? geburtstag = null;
        foreach (var raw in lines)
        {
            if (!TrySplit(raw, out var prop, out var rawValue)) { continue; }
            var propParts = prop.Split(';');  // Die Basis-Eigenschaft (vor dem ersten Semikolon) extrahieren
            var baseProp = propParts[0];
            var value = rawValue;
            if (prop.Contains("QUOTED-PRINTABLE", StringComparison.OrdinalIgnoreCase)) { value = DecodeQuotedPrintable(value); }  // vCard 2.1 Kompatibilität: löst Umlaute und Zeilenumbrüche
            switch (baseProp)
            {
                case "N":
                    var nParts = SplitValue(value, ';', 5);
                    nachname = UnEsc(nParts[0]);
                    vorname = UnEsc(nParts[1]);
                    zwischen = UnEsc(nParts[2]);
                    praefix = UnEsc(nParts[3]);
                    suffix = UnEsc(nParts[4]);
                    break;
                case "FN":  // Nur als Fallback, wenn N fehlt oder leer war
                    if (string.IsNullOrWhiteSpace(nachname) && string.IsNullOrWhiteSpace(vorname))
                    {
                        var fnParts = UnEsc(value).Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);  // Versuche "Vorname Nachname" zu trennen
                        if (fnParts.Length == 2) { vorname = fnParts[0]; nachname = fnParts[1]; }
                        else if (fnParts.Length == 1) { nachname = fnParts[0]; }
                    }
                    break;
                case "NICKNAME":
                    nickname = UnEsc(value);
                    break;
                case "X-ANREDE":
                    anrede = UnEsc(value);
                    break;
                case "ORG":
                    unternehmen = UnEsc(value.Split(';')[0]);
                    break;
                case "TITLE":
                    position = UnEsc(value);
                    break;
                case "ADR":
                    var adrParts = SplitValue(value, ';', 7);
                    postfach ??= UnEsc(adrParts[0]);  // Standardmäßig nehmen wir die erste gefundene Adresse, es sei denn, wir differenzieren nach WORK/HOME
                    strasse ??= UnEsc(adrParts[2]);
                    ort ??= UnEsc(adrParts[3]);
                    plz ??= UnEsc(adrParts[5]);
                    land ??= UnEsc(adrParts[6]);
                    break;
                case "TEL":
                    var phone = UnEsc(value);
                    if (string.IsNullOrWhiteSpace(phone)) { break; }
                    if (prop.Contains("CELL", StringComparison.OrdinalIgnoreCase) || prop.Contains("MOBILE", StringComparison.OrdinalIgnoreCase)) { mobil ??= phone; }
                    else if (prop.Contains("FAX", StringComparison.OrdinalIgnoreCase)) { fax ??= phone; }
                    else
                    {
                        phoneCount++;
                        if (phoneCount == 1) { tel1 = phone; }
                        else if (phoneCount == 2) { tel2 = phone; }
                    }
                    break;
                case "EMAIL":
                    var mail = UnEsc(value);
                    if (string.IsNullOrWhiteSpace(mail)) { break; }
                    mailCount++;
                    if (mailCount == 1) { mail1 = mail; }
                    else if (mailCount == 2) { mail2 = mail; }
                    break;
                case "URL":
                    internet = UnEsc(value);
                    break;
                case "BDAY":
                    if (TryParseBday(value, out var bday)) { geburtstag = bday; }
                    break;
                case "NOTE":
                    notizen = UnEscNote(value);
                    break;
                case "CATEGORIES":
                    gruppenNamen = [.. value.Split(',').Select(s => UnEsc(s.Trim())).Where(s => !string.IsNullOrWhiteSpace(s))];
                    break;
                case "X-BETREFF":
                    betreff = UnEsc(value);
                    break;
                case "X-GRUSSFORMEL":
                    gruss = UnEsc(value);
                    break;
                case "X-SCHLUSSFORMEL":
                    schluss = UnEsc(value);
                    break;
                case "PHOTO":
                    if (prop.Contains("VALUE=URI", StringComparison.OrdinalIgnoreCase) || value.StartsWith("http", StringComparison.OrdinalIgnoreCase)) { photoUrl = value.Trim(); }
                    else
                    {
                        var b64 = value.Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("\n", "");  // Entfernt konsequent alle Whitespaces und Zeilenumbrüche
                        if (b64.Length % 4 != 0) { b64 = b64.PadRight(b64.Length + (4 - b64.Length % 4), '='); }  // Base64-Strings auf ein Vielfaches von 4 auffüllen (Fehlendes Padding reparieren)
                        try { fotoBytes = Convert.FromBase64String(b64); }
                        catch { }  // fehlerhafte Base64-Strings stumm ignorieren 
                    }
                    break;
            }
        }
        if (asContact)
        {
            var c = new Contact
            {
                Anrede = anrede,
                Praefix = praefix,
                Nachname = nachname,
                Vorname = vorname,
                Zwischenname = zwischen,
                Nickname = nickname,
                Suffix = suffix,
                Unternehmen = unternehmen,
                Position = position,
                Strasse = strasse,
                PLZ = plz,
                Ort = ort,
                Postfach = postfach,
                Land = land,
                Telefon1 = tel1,
                Telefon2 = tel2,
                Mobil = mobil,
                Fax = fax,
                Mail1 = mail1,
                Mail2 = mail2,
                Internet = internet,
                Geburtstag = geburtstag,
                Notizen = notizen,
                Betreff = betreff,
                Grussformel = gruss,
                Schlussformel = schluss,
                PhotoUrl = photoUrl,
                GroupNames = gruppenNamen
            };
            return new VCardImportResult { Contact = c, Adresse = null, GruppenNamen = gruppenNamen, FotoBytes = fotoBytes }; // FotoBytes mitgeben!
        }
        else
        {
            var a = new Adresse
            {
                Anrede = anrede,
                Praefix = praefix,
                Nachname = nachname,
                Vorname = vorname,
                Zwischenname = zwischen,
                Nickname = nickname,
                Suffix = suffix,
                Unternehmen = unternehmen,
                Position = position,
                Strasse = strasse,
                PLZ = plz,
                Ort = ort,
                Postfach = postfach,
                Land = land,
                Telefon1 = tel1,
                Telefon2 = tel2,
                Mobil = mobil,
                Fax = fax,
                Mail1 = mail1,
                Mail2 = mail2,
                Internet = internet,
                Geburtstag = geburtstag,
                Notizen = notizen,
                Betreff = betreff,
                Grussformel = gruss,
                Schlussformel = schluss,
            };
            return new VCardImportResult { Adresse = a, Contact = null, GruppenNamen = gruppenNamen, FotoBytes = fotoBytes };
        }
    }

    private static string DecodeQuotedPrintable(string input)
    { //Decodiert Quoted-Printable Strings (z.B. "=D6" zu "Ö", "=0D=0A" zu Zeilenumbruch). Nutzt standardmäßig Latin1 (ISO-8859-1), da vCard 2.1 historisch oft in diesem Zeichensatz gespeichert wurde.
        if (string.IsNullOrEmpty(input)) { return string.Empty; }
        var bytes = new List<byte>();
        for (var i = 0; i < input.Length; i++)
        {
            if (input[i] == '=' && i + 2 < input.Length)
            {
                var hex = input.Substring(i + 1, 2);
                if (byte.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
                {
                    bytes.Add(b);
                    i += 2;
                }
                bytes.Add(input[i] <= 0xFF ? (byte)input[i] : (byte)'?');  // Bei echtem ASCII-Zeichen das Byte direkt übernehmen, sonst Fragezeichen (Daten sind hier ohnehin defekt)
            }
            else { bytes.Add((byte)input[i]); }
        }
        return Encoding.Latin1.GetString([.. bytes]);  // Encoding.Latin1 entspricht ISO-8859-1 und ist in .NET Core/10 standardmäßig verfügbar.
    }

    private static void Begin(StringBuilder sb)
    {
        sb.AppendLine("BEGIN:VCARD");
        sb.AppendLine("VERSION:3.0");
    }

    private static void End(StringBuilder sb) => sb.AppendLine("END:VCARD");

    private static string BuildFN(string? vorname, string? zwischen, string? nachname, string? firma)
    {
        var parts = new[] { vorname, zwischen, nachname }.Where(s => !string.IsNullOrWhiteSpace(s));
        var name = string.Join(" ", parts).Trim();
        return string.IsNullOrWhiteSpace(name) ? (firma ?? string.Empty) : name;
    }

    private static string Esc(string? s)
    {
        if (string.IsNullOrEmpty(s)) { return string.Empty; }
        return s.Replace("\\", "\\\\").Replace(",", "\\,").Replace(";", "\\;");  // RFC 6350 Escaping: \, \\, \;, Newlines bleiben (außer in NOTE)
    }

    private static string EscNote(string? s)
    {
        if (string.IsNullOrEmpty(s)) { return string.Empty; }
        return s.Replace("\\", "\\\\").Replace(",", "\\,").Replace(";", "\\;").Replace("\r\n", "\\n").Replace("\n", "\\n").Replace("\r", "\\n");
    }

    private static string UnEsc(string? s)
    {
        if (string.IsNullOrEmpty(s)) { return string.Empty; }
        return s.Replace("\\\\", "\x00").Replace("\\;", ";").Replace("\\,", ",").Replace("\x00", "\\").Trim();
    }

    private static string UnEscNote(string? s)
    {
        if (string.IsNullOrEmpty(s)) { return string.Empty; }
        return s.Replace("\\\\", "\x00").Replace("\\n", "\r\n").Replace("\\N", "\r\n").Replace("\\;", ";").Replace("\\,", ",").Replace("\x00", "\\").Trim();
    }

    private static bool HasAny(params string?[] values) => values.Any(v => !string.IsNullOrWhiteSpace(v));

    private static string[] SplitValue(string value, char sep, int count)
    {
        var parts = value.Split(sep);
        var result = new string[count];
        for (var i = 0; i < count; i++) { result[i] = i < parts.Length ? parts[i] : string.Empty; }
        return result;
    }

    private static bool TrySplit(string line, out string prop, out string value)
    { //Trennt Eigenschaft (mit Parametern) und Wert – z. B. "TEL;TYPE=CELL:012345" → ("TEL;TYPE=CELL", "012345"). Gibt false für Steuerzeilen (BEGIN/END/VERSION …) und Leerzeilen zurück.
        prop = value = string.Empty;
        if (string.IsNullOrWhiteSpace(line)) { return false; }
        var colonIdx = line.IndexOf(':');
        if (colonIdx <= 0) { return false; }
        prop = line[..colonIdx].Trim().ToUpperInvariant();
        value = line[(colonIdx + 1)..].Trim();  // ← .Trim() entfernt Whitespace am ENDE
        return prop is not ("BEGIN" or "END" or "VERSION" or "PRODID" or "REV" or "UID");

    }

    private static IEnumerable<string> UnfoldLines(string[] rawLines)
    {
        var current = new StringBuilder();
        var inPhoto = false;

        foreach (var line in rawLines)
        {
            if (string.IsNullOrWhiteSpace(line)) { continue; }

            var startsWithFold = line[0] == ' ' || line[0] == '\t';
            var prevEndsWithQP = !inPhoto && current.Length > 0 && current[^1] == '='; // ← !inPhoto

            if (startsWithFold && current.Length > 0) { current.Append(line.TrimStart()); }
            else if (prevEndsWithQP)
            {
                current.Remove(current.Length - 1, 1);
                current.Append(line);
            }
            else if (inPhoto && !line.StartsWith("END:", StringComparison.OrdinalIgnoreCase) && !line.Contains(':')) { current.Append(line); }
            else
            {
                if (current.Length > 0) { yield return current.ToString(); }
                current.Clear();
                current.Append(line);
                inPhoto = line.StartsWith("PHOTO", StringComparison.OrdinalIgnoreCase);
            }
        }
        if (current.Length > 0) { yield return current.ToString(); }
    }

    private static IEnumerable<string> Chunks(string s, int size)
    {
        for (var i = 0; i < s.Length; i += size) { yield return s.Substring(i, Math.Min(size, s.Length - i)); }
    }

    private static bool TryParseBday(string value, out DateOnly result)
    {
        result = default;
        if (DateOnly.TryParseExact(value, ["yyyy-MM-dd", "yyyyMMdd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out result)) { return true; }
        if (value.StartsWith("--") && value.Length >= 5)  // Jahrloser Geburtstag nach RFC: "--MM-dd" oder "--MMdd"
        {
            var withoutDash = "2000" + value[2..].Replace("-", "");
            if (DateOnly.TryParseExact(withoutDash, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var partial))
            {
                result = new DateOnly(1, partial.Month, partial.Day);
                return true;
            }
        }
        return false;
    }

    private static string DetectImageType(byte[] data)
    {
        if (data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8) { return "JPEG"; }
        if (data.Length >= 4 && data[0] == 0x89 && data[1] == 0x50) { return "PNG"; }
        if (data.Length >= 2 && data[0] == 0x42 && data[1] == 0x4D) { return "BMP"; }
        return "JPEG"; // Fallback
    }
}