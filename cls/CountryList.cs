namespace Adressen.cls;

/// <summary>
/// Statische Liste der wichtigsten Ländernamen (Europa vollständig, plus weitere bedeutende Staaten weltweit),
/// z. B. als AutoCompleteCustomSource für das Feld "Land" (cbLand in FrmAdressen).
///
/// Deutsche Schreibweise (Absendersprache) gemäß DIN 5008 / Deutsche Post-Empfehlung für Auslandsanschriften,
/// ohne Länderkennzeichen wie "CH-" oder "F-" (diese wurden 1999 offiziell abgeschafft, da sie zu
/// Zustellproblemen führten). Normale Groß-/Kleinschreibung, damit die Liste im Programm nicht wie Versalien
/// wirkt — die für den Umschlagdruck empfohlene Schreibweise in Großbuchstaben lässt sich separat über die
/// Option ckbLandGROSS (AppSettings.RecipientCountryUpper) aktivieren.
/// </summary>
public static class CountryList
{
    public static readonly string[] Names =
    [
        "Ägypten", "Albanien", "Algerien", "Andorra", "Argentinien", "Australien",
        "Belarus", "Belgien", "Bosnien und Herzegowina", "Brasilien", "Bulgarien",
        "Chile", "China",
        "Dänemark", "Deutschland",
        "Estland",
        "Finnland", "Frankreich",
        "Griechenland", "Großbritannien",
        "Indien", "Indonesien", "Irland", "Island", "Israel", "Italien",
        "Japan",
        "Kanada", "Kolumbien", "Kosovo", "Kroatien",
        "Lettland", "Liechtenstein", "Litauen", "Luxemburg",
        "Malta", "Marokko", "Mexiko", "Moldau", "Monaco", "Montenegro",
        "Neuseeland", "Niederlande", "Nordmazedonien", "Norwegen",
        "Österreich",
        "Pakistan", "Peru", "Polen", "Portugal",
        "Rumänien", "Russland",
        "Saudi-Arabien", "Schweden", "Schweiz", "Serbien", "Singapur", "Slowakei", "Slowenien", "Spanien", "Südafrika", "Südkorea",
        "Thailand", "Tschechien", "Türkei",
        "Ukraine", "Ungarn",
        "Vatikanstadt", "Vereinigte Arabische Emirate", "USA", "Vereinigtes Königreich", "Vietnam",
        "Zypern",
    ];
}
