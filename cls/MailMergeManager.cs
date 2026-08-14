using System.IO.Compression;
using System.Reflection;
using System.Text;

namespace Adressen.cls;

/// <summary>
/// Erzeugt aus den aktuell angezeigten Datensätzen (Adresse oder Contact) eine
/// CSV-Datenquelle für den Serienbriefdruck mit Microsoft Word bzw. LibreOffice Writer.
/// Die Spaltennamen entsprechen exakt den Textmarken-Namen aus FillWordProcessingDictionary,
/// damit dieselben Feldnamen wie in den Dokumenten verwendet werden können.
/// </summary>
internal static class MailMergeManager
{
    /// <summary>Dateiname der erzeugten CSV-Datenquelle (muss zum schema.ini-Abschnitt passen).</summary>
    internal const string CsvFileName = "Serienbrief.csv";

    /// <summary>Dateiname der LibreOffice-Datenquelle (.odb) mit fertigen Verbindungseinstellungen.</summary>
    internal const string OdbFileName = "Serienbrief.odb";

    /// <summary>Reihenfolge und vollständige Liste der Spalten in der CSV-Datenquelle.</summary>
    internal static readonly string[] FieldOrder =
    [
        "Anrede", "Praefix", "Vorname", "Zwischenname", "Zwischenname_initial", "Nickname", "Nachname", "Suffix",
        "Unternehmen", "Position",
        "Anr_Prf_Vor_Nach",
        "Anr_Prf_Vor_Zw_Nach",
        "Anr_Prf_Vor_ZwI_Nach",
        "Prf_Vor_Nach",
        "Prf_Vor_Zw_Nach",
        "Prf_Vor_ZwI_Nach",
        "Vor_Nach",
        "Vor_Zw_Nach",
        "Vor_ZwI_Nach",
        "Adresse", "Postfach", "Postfach_sonst_Adresse",
        "PLZ", "Ort", "PLZ_Ort", "Land", "Land_Gross",
        "Betreff", "Grussformel", "Schlussformel",
        "Telefon1", "Telefon2", "Mail1", "Mail2", "Mobil", "Fax", "Internet",
    ];

    // Cache der PropertyInfos pro Typ, damit die Reflection bei vielen Datensätzen schnell bleibt.
    private static readonly Dictionary<Type, Dictionary<string, PropertyInfo>> _propCache = [];

    /// <summary>
    /// Schreibt alle übergebenen Datensätze als CSV (UTF-8 ohne BOM, Semikolon-getrennt, gequotet)
    /// und legt daneben eine schema.ini an, damit der Word-Texttreiber Trennzeichen, Kopfzeile,
    /// Zeichensatz und Spaltentypen (alle Text → keine verlorenen führenden Nullen in der PLZ) kennt.
    /// </summary>
    /// <returns>Anzahl der geschriebenen Datensätze.</returns>
    internal static int WriteCsv(string csvPath, IEnumerable<object> entities)
    {
        var dir = Path.GetDirectoryName(csvPath);
        if (!string.IsNullOrEmpty(dir)) { Directory.CreateDirectory(dir); }

        var count = 0;
        using (var writer = new StreamWriter(csvPath, append: false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
        {
            writer.WriteLine(string.Join(";", FieldOrder.Select(EscapeCsv)));
            foreach (var entity in entities)
            {
                var fields = BuildFields(entity);
                writer.WriteLine(string.Join(";", FieldOrder.Select(key => EscapeCsv(fields.TryGetValue(key, out var v) ? v : string.Empty))));
                count++;
            }
        }
        WriteSchemaIni(csvPath);
        return count;
    }

    /// <summary>
    /// Baut für einen einzelnen Datensatz das Spalten-Dictionary auf – Rohfelder plus
    /// zusammengesetzte Felder. Die Logik spiegelt FillWordProcessingDictionary in FrmAdressen.cs.
    /// Funktioniert per Reflection für Adresse und Contact (identische Property-Namen).
    /// </summary>
    internal static Dictionary<string, string> BuildFields(object entity)
    {
        var anrede = Raw(entity, "Anrede");
        var praefix = Raw(entity, "Praefix");
        var vorname = Raw(entity, "Vorname");
        var zwischen = Raw(entity, "Zwischenname");
        var nachname = Raw(entity, "Nachname");
        var nickname = Raw(entity, "Nickname");
        var suffix = Raw(entity, "Suffix");
        var firma = Raw(entity, "Unternehmen");
        var position = Raw(entity, "Position");
        var strasse = Raw(entity, "Strasse");
        var postfach = Raw(entity, "Postfach");
        var plz = Raw(entity, "PLZ");
        var ort = Raw(entity, "Ort");
        var land = Raw(entity, "Land");
        var betreff = Raw(entity, "Betreff");
        var gruss = Raw(entity, "Grussformel");
        var schluss = Raw(entity, "Schlussformel");
        var geburtstag = Raw(entity, "Geburtstag");
        var mail1 = Raw(entity, "Mail1");
        var mail2 = Raw(entity, "Mail2");
        var tel1 = Raw(entity, "Telefon1");
        var tel2 = Raw(entity, "Telefon2");
        var mobil = Raw(entity, "Mobil");
        var fax = Raw(entity, "Fax");
        var internet = Raw(entity, "Internet");
        //var notizen = Raw(entity, "Notizen");
        var zwischenInitial = string.IsNullOrEmpty(zwischen) ? string.Empty : $"{zwischen[0]}.";

        static string Join(params string?[] parts) => string.Join(" ", parts.Where(static s => !string.IsNullOrWhiteSpace(s)));

        return new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["Anrede"] = anrede,
            ["Praefix"] = praefix,
            ["Vorname"] = vorname,
            ["Zwischenname"] = zwischen,
            ["Zwischenname_initial"] = zwischenInitial,
            ["Nickname"] = nickname,
            ["Nachname"] = nachname,
            ["Suffix"] = suffix,
            ["Unternehmen"] = firma,
            ["Position"] = position,
            ["Anr_Prf_Vor_Nach"] = Join(anrede, praefix, vorname, nachname),
            ["Anr_Prf_Vor_Zw_Nach"] = Join(anrede, praefix, vorname, zwischen, nachname),
            ["Anr_Prf_Vor_ZwI_Nach"] = Join(anrede, praefix, vorname, zwischenInitial, nachname),
            ["Prf_Vor_Nach"] = Join(praefix, vorname, nachname),
            ["Prf_Vor_Zw_Nach"] = Join(praefix, vorname, zwischen, nachname),
            ["Prf_Vor_ZwI_Nach"] = Join(praefix, vorname, zwischenInitial, nachname),
            ["Vor_Nach"] = Join(vorname, nachname),
            ["Vor_Zw_Nach"] = Join(vorname, zwischen, nachname),
            ["Vor_ZwI_Nach"] = Join(vorname, zwischenInitial, nachname),
            ["Adresse"] = strasse,
            ["Postfach"] = postfach,
            ["Postfach_sonst_Adresse"] = string.IsNullOrEmpty(postfach) ? strasse : $"Postfach {postfach}",
            ["PLZ"] = plz,
            ["Ort"] = ort,
            ["PLZ_Ort"] = $"{plz} {ort}".Trim(),
            ["Land"] = land,
            ["Land_Gross"] = land.ToUpperInvariant(),
            ["Betreff"] = betreff,
            ["Grussformel"] = gruss,
            ["Schlussformel"] = schluss,
            ["Geburtstag"] = geburtstag,
            ["Mail1"] = mail1,
            ["Mail2"] = mail2,
            ["Telefon1"] = tel1,
            ["Telefon2"] = tel2,
            ["Mobil"] = mobil,
            ["Fax"] = fax,
            ["Internet"] = internet,
            //["Notizen"] = notizen,
        };
    }

    /// <summary>Liest eine Property per Reflection aus und liefert sie als getrimmten String (DateOnly → dd.MM.yyyy).</summary>
    private static string Raw(object entity, string propertyName)
    {
        var type = entity.GetType();
        if (!_propCache.TryGetValue(type, out var map))
        {
            map = type.GetProperties().Where(static p => p.CanRead).GroupBy(static p => p.Name).ToDictionary(static g => g.Key, static g => g.First());
            _propCache[type] = map;
        }
        if (!map.TryGetValue(propertyName, out var pi)) { return string.Empty; }
        var value = pi.GetValue(entity);
        return value switch
        {
            null => string.Empty,
            DateOnly d => d.ToString("dd.MM.yyyy"),
            DateTime dt => dt.ToString("dd.MM.yyyy"),
            _ => value.ToString()?.Trim() ?? string.Empty,
        };
    }

    /// <summary>
    /// Quotet ein CSV-Feld. Zeilenumbrüche werden durch Leerzeichen ersetzt und – ganz wichtig – ein
    /// eingebettetes Semikolon (= das Trennzeichen) wird durch ein Komma ersetzt. Andernfalls würde der
    /// Word-Seriendruck die Zeile in zu viele Felder zerlegen ("Datensatz X enthält zu viele Datenfelder").
    /// </summary>
    private static string EscapeCsv(string value)
    {
        var clean = value
            .Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ')
            .Replace(';', ','); // Trennzeichen im Wert neutralisieren
        return $"\"{clean.Replace("\"", "\"\"")}\"";
    }

    /// <summary>
    /// Legt eine schema.ini neben der CSV an. Sie weist den Microsoft-Texttreiber (vom Word-Seriendruck
    /// genutzt) an: Semikolon als Trennzeichen, erste Zeile = Kopfzeile, UTF-8, alle Spalten als Text.
    /// </summary>
    private static void WriteSchemaIni(string csvPath)
    {
        var dir = Path.GetDirectoryName(csvPath);
        if (string.IsNullOrEmpty(dir)) { return; }
        var fileName = Path.GetFileName(csvPath);

        var sb = new StringBuilder();
        sb.AppendLine($"[{fileName}]");
        sb.AppendLine("ColNameHeader=True");
        sb.AppendLine("Format=Delimited(;)");
        sb.AppendLine("CharacterSet=65001"); // UTF-8
        sb.AppendLine("MaxScanRows=0");
        for (var i = 0; i < FieldOrder.Length; i++)
        {
            // Alle Spalten als Text deklarieren → führende Nullen (PLZ) und Datumsformate bleiben erhalten.
            sb.AppendLine($"Col{i + 1}={FieldOrder[i]} Text Width 4000");
        }
        File.WriteAllText(Path.Combine(dir, "schema.ini"), sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    /// <summary>
    /// Erzeugt eine LibreOffice-Datenquelle (.odb) für die CSV. Anders als Word liest LibreOffice keine
    /// schema.ini; die Verbindungseinstellungen (Feldtrenner »;«, Texttrenner »"«, Zeichensatz UTF-8,
    /// Kopfzeile) werden deshalb direkt in der .odb hinterlegt. Der Anwender muss sie dann nicht mehr
    /// von Hand korrigieren, sondern bindet die .odb im Seriendruck-Assistenten nur noch ein.
    /// Die .odb verweist auf den Ordner; die CSV-Datei darin ist die Tabelle (Name = Dateiname ohne ».csv«).
    /// </summary>
    internal static void WriteOdb(string odbPath, string csvFolder)
    {
        var folder = csvFolder.EndsWith(Path.DirectorySeparatorChar) ? csvFolder : csvFolder + Path.DirectorySeparatorChar;
        var hrefEsc = XmlAttr(new Uri(folder).AbsoluteUri);  // z.B. file:///C:/Users/Me/Downloads/Adressen/

        var content =
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:ooo="http://openoffice.org/2004/office" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
            xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
            xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
            xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:rpt="http://openoffice.org/2005/report"
            xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
            xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
            xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
            xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc"
            xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xforms="http://www.w3.org/2002/xforms"
            xmlns:tableooo="http://openoffice.org/2009/table" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"
            xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:xhtml="http://www.w3.org/1999/xhtml"
            xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"
            xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
            xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0"
            xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xsd="http://www.w3.org/2001/XMLSchema"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:grddl="http://www.w3.org/2003/g/data-view#"
            xmlns:css3t="http://www.w3.org/TR/css3-text/" xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" office:version="1.4">
            <office:scripts/><office:font-face-decls/><office:automatic-styles/>
            <office:body><office:database>
            <db:data-source>
            <db:connection-data>
            <db:database-description>
            <db:file-based-database xlink:href="{hrefEsc}" db:media-type="text/csv" db:extension="csv"/>
            </db:database-description>
            <db:login db:is-password-required="false"/>
            </db:connection-data>
            <db:driver-settings db:system-driver-settings="" db:base-dn="">
            <db:delimiter db:field=";" db:string="&quot;" db:decimal="." db:thousand=","/>
            <db:font-charset db:encoding="UTF-8"/>
            </db:driver-settings>
            <db:application-connection-settings db:is-table-name-length-limited="false" db:append-table-alias-name="false" db:max-row-count="100">
            <db:table-filter>
            <db:table-include-filter>
            <db:table-filter-pattern>%</db:table-filter-pattern>
            </db:table-include-filter>
            </db:table-filter>
            <db:data-source-settings>
            <db:data-source-setting db:data-source-setting-is-list="false" db:data-source-setting-name="Extension" db:data-source-setting-type="string"><db:data-source-setting-value>csv</db:data-source-setting-value></db:data-source-setting>
            </db:data-source-settings>
            </db:application-connection-settings>
            </db:data-source>
            <db:queries>
            <db:query db:name="Adressen" db:command="SELECT &quot;Anrede&quot; AS &quot;Anrede&quot;, &quot;Praefix&quot; AS &quot;Praefix&quot;, &quot;Vorname&quot; AS &quot;Vorname&quot;, &quot;Zwischenname&quot; AS &quot;Zwischenname&quot;, &quot;Zwischenname_initial&quot; AS &quot;Zwischenname_initial&quot;, &quot;Nickname&quot; AS &quot;Nickname&quot;, &quot;Nachname&quot; AS &quot;Name&quot;, &quot;Suffix&quot; AS &quot;Suffix&quot;, &quot;Unternehmen&quot; AS &quot;Firmenname&quot;, &quot;Position&quot; AS &quot;Position&quot;, &quot;Adresse&quot; AS &quot;Adresszeile 1&quot;, &quot;Postfach&quot; AS &quot;Postfach&quot;, &quot;PLZ&quot; AS &quot;PLZ&quot;, &quot;Ort&quot; AS &quot;Stadt&quot;, &quot;Land&quot; AS &quot;Land&quot;, &quot;Land_Gross&quot; AS &quot;Land_Gross&quot;, &quot;Betreff&quot; AS &quot;Betreff&quot;, &quot;Grussformel&quot; AS &quot;Grussformel&quot;, &quot;Schlussformel&quot; AS &quot;Schlussformel&quot;, &quot;Telefon1&quot; AS &quot;Telefon geschäftlich&quot;, &quot;Telefon2&quot; AS &quot;Telefon privat&quot;, &quot;Mail1&quot; AS &quot;E-Mail-Adresse&quot;, &quot;Internet&quot; AS &quot;Webseite&quot; FROM &quot;Serienbrief&quot;" db:escape-processing="false"/>
            </db:queries>
            <db:table-representations>
            <db:table-representation db:name="Serienbrief"/>
            </db:table-representations>
            </office:database></office:body></office:document-content>
            """;

        const string settings =
            """
            <?xml version="1.0" encoding="UTF-8"?>
            <office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:svg="http://www.w3.org/2000/svg" xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" office:version="1.4"/>
            """;

        const string manifest =
            """
            <?xml version="1.0" encoding="UTF-8"?>
            <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.4" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">
             <manifest:file-entry manifest:full-path="/" manifest:version="1.4" manifest:media-type="application/vnd.oasis.opendocument.base"/>
             <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
             <manifest:file-entry manifest:full-path="Configurations2/" manifest:media-type="application/vnd.sun.xml.ui.configuration"/>
             <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
            </manifest:manifest>
            """;

        if (File.Exists(odbPath)) { File.Delete(odbPath); }
        using var zip = ZipFile.Open(odbPath, ZipArchiveMode.Create);

        // mimetype als erster Eintrag und unkomprimiert (OpenDocument-Vorgabe).
        var mimeEntry = zip.CreateEntry("mimetype", CompressionLevel.NoCompression);
        using (var s = mimeEntry.Open())
        {
            var bytes = Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.base");
            s.Write(bytes, 0, bytes.Length);
        }
        WriteZipText(zip, "content.xml", content);
        zip.CreateEntry("Configurations2/", CompressionLevel.NoCompression); // leere Ordner wie im LibreOffice-Original
        zip.CreateEntry("reports/", CompressionLevel.NoCompression);
        zip.CreateEntry("forms/", CompressionLevel.NoCompression);
        WriteZipText(zip, "settings.xml", settings);
        WriteZipText(zip, "META-INF/manifest.xml", manifest);
    }

    private static void WriteZipText(ZipArchive zip, string entryName, string text)
    {
        var entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
        using var s = entry.Open();
        var bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(text);
        s.Write(bytes, 0, bytes.Length);
    }

    private static string XmlAttr(string value) => value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
}
