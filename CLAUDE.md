# Adressen

„Adressen & Kontakte" — WinForms-Adressbuch (.NET 10, ausschließlich x64) von Wilhelm Happe („ophthalmos"): lokale Adressen (SQLite über EF Core) und Google-Kontakte (People API) in einer Oberfläche; Übertragung in Word- und LibreOffice-Dokumente über Textmarken, Briefumschlagdruck, Geburtstagserinnerung, Fritz!Box-Anrufmonitor. Benutzer werden in allen Programmtexten **geduzt**. Der Projektordner heißt `Adressen11` (der Nachbarordner `Adressen10` ist die Vorgängerversion); Assembly und sichtbarer Name sind „Adressen".

## Bauen

- `dotnet build` auf `Adressen.csproj` (Solution: `Adressen.slnx`), Plattform nur x64.
- Build-Voraussetzungen: installiertes Word (COMReference `Microsoft.Office.Interop.Word`) und 7-Zip (`C:\Program Files\7-Zip\7z.exe`) — das PostBuild-Ereignis packt `frm\` und `cls\` als Wochentags-Backup nach `_Backups\`.
- `client_secret.json` (Google-OAuth-Client) ist bewusst git-ignoriert — nie committen oder nach außen geben.
- Installer: `Adressen.iss` mit **Inno Setup 7** kompilieren (`C:\Program Files\Inno Setup 7\ISCC.exe` — nicht die parallel installierte v6). Versionsnummer an zwei Stellen pflegen: csproj (`AssemblyVersion`/`FileVersion`) und iss (`MyAppVersion`).
- Veröffentlichung: `AdressenSetup.exe` auf www.netradio.info (macht Wilhelm selbst), danach WinGet über `winget-release.ps1` (wingetcreate, Paket `WilhelmHappe.Adressen`). Gepusht wird nur von Wilhelm (GitHub-Stand = veröffentlichter Release-Stand).

## Aufbau

- `frm\FrmAdressen` — Hauptfenster; daneben u. a. FrmAdvSearch (erweiterte Suche), FrmBirthdays, FrmImportCsv, FrmPrintSetting/FrmSinglePrintPreview (Druck), FrmCopyScheme, FrmGroupsEdit/-Filter/-Rename (Gruppen), FrmProgSettings, FrmSplashScreen.
- `cls\` — AdressenDbContext + DatabaseMigrator (SQLite; **COLLATE NOCASE**, betrifft nur Sortieren/Finden — siehe `Hinweise.txt`), GooglePeopleManager (People API/OAuth), WordManager + MailMergeManager (Word-Interop), VCardService, FritzCallMonitor, HtmlPrintService, SettingsManager, Utilities (u. a. `HelpMsgTaskDlg` — Vorlage der Über-Dialoge auch in Wilhelms anderen Projekten), eigene Controls (TagControl, YearSlider, IPv4AddressControl, PaddedTextBox, FlickerFreePrintPreviewControl).
- `LibreOffice\`, `Papierkorb\`, `_Backups\` sind vom Build ausgenommen (Altstände/Ablage).

## Konventionen

- **Designer-Regel:** Sämtlicher Code, der im Visual-Studio-Designer-Inspector stehen könnte, gehört in die `.Designer.cs` (`InitializeComponent`): keine Forms komplett in Code bauen, kein Control-Aufbau, keine Property-Zuweisungen oder Event-Verdrahtungen (`FormClosing +=` usw.) im Form-Code, keine Lambdas als Event-Handler für Designer-Controls — benannte Methoden verwenden. Im Form-Code bleibt nur, was der Designer nicht kann: Laufzeitdaten, dynamisch gerenderte Bilder, Renderer. Wartbarkeit über den VS-Designer hat Priorität.
- Codestil: `var` statt expliziter Typen; `Nullable` ist aktiviert.
- Git: mehrzeilige Commit-Botschaften mit Umlauten scheitern in PowerShell an `git -m` — stattdessen `git commit -F <datei>` (Datei ohne BOM schreiben, z. B. `[IO.File]::WriteAllText`).
