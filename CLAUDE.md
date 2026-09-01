# Adressen

WinForms-Programm (.NET 10) von Wilhelm Happe („ophthalmos"). Benutzer werden in allen Programmtexten **geduzt**.

## Konventionen

- **Designer-Regel:** Sämtlicher Code, der im Visual-Studio-Designer-Inspector stehen könnte, gehört in die `.Designer.cs` (`InitializeComponent`): keine Forms komplett in Code bauen, kein Control-Aufbau, keine Property-Zuweisungen oder Event-Verdrahtungen (`FormClosing +=` usw.) im Form-Code, keine Lambdas als Event-Handler für Designer-Controls — benannte Methoden verwenden. Im Form-Code bleibt nur, was der Designer nicht kann: Laufzeitdaten, dynamisch gerenderte Bilder, Renderer. Wartbarkeit über den VS-Designer hat Priorität.
- Codestil: `var` statt expliziter Typen.
