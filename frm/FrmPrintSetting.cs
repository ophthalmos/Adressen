using System.Drawing.Printing;
using Adressen.cls;

namespace Adressen;

public partial class FrmPrintSetting : Form
{
    private readonly double _zoom = 0.50F;
    private readonly AppSettings _settings;
    private readonly BindingSource _bindingSource; // Der Vermittler für das DataBinding
    private readonly Dictionary<string, string> _recipientDict;

    internal FrmPrintSetting(AppSettings settings, Dictionary<string, string> recipientDict)
    {
        InitializeComponent();
        _settings = settings;
        _recipientDict = recipientDict;
        _bindingSource = new BindingSource { DataSource = _settings };
        foreach (TabPage tabPage in tabControl.TabPages)
        {
            tabPage.BackColor = _settings.ColorScheme switch
            {
                "blue" => SystemColors.InactiveBorder,
                "pale" => SystemColors.ControlLightLight,
                "dark" => SystemColors.AppWorkspace,
                _ => SystemColors.ButtonFace,
            };
        }
        cbFont.Items.AddRange([.. FontFamily.Families.Select(f => f.Name)]);  // muss vor dem Binding passieren
        foreach (string s in PrinterSettings.InstalledPrinters) { cbPrinter.Items.Add(s); }
        InitializeDataBindings();
        UpdateUiState();  // UI Logik (Enable/Disable) initial anstoßen
        InitializePrinterSelection();  // Drucker initialisieren (für Papierformate etc.)
    }

    private void InitializeDataBindings()
    {
        // Hilfsmethode um Schreibarbeit zu sparen
        // OnPropertyChanged sorgt dafür, dass Änderungen sofort im Objekt landen (wichtig für Preview)
        void Bind(Control control, string propertyName, string dataMember) { control.DataBindings.Add(propertyName, _bindingSource, dataMember, true, DataSourceUpdateMode.OnPropertyChanged); }

        // --- Tab 1: Format ---
        Bind(cbPrinter, "Text", nameof(AppSettings.PrintDevice));
        Bind(cbSources, "Text", nameof(AppSettings.PrintSource));
        Bind(cbPapersize, "Text", nameof(AppSettings.PrintFormat));

        // RadioButtons: Wir binden nur den "Haupt"-Button. Der andere ergibt sich logisch.
        Bind(rbLandscape, "Checked", nameof(AppSettings.PrintLandscape));
        // Hinweis: rbPortrait muss nicht gebunden werden, da sie im gleichen Container sind und sich gegenseitig ausschalten.
        // Wir setzen rbPortrait nur initial, falls Landscape false ist, passiert aber automatisch durch WinForms Logik oft schon.
        // Um sicherzugehen, setzen wir einen Event-Handler im Designer oder Code, der rbPortrait toggelt, 
        // aber das Binding auf Landscape reicht meistens als "Master".

        // --- Tab 2: Schriftarten ---
        Bind(cbFont, "Text", nameof(AppSettings.PrintFont));
        Bind(cbFontsizeSender, "Text", nameof(AppSettings.SenderFontsize));     // Int zu String Konvertierung macht Binding automatisch
        Bind(cbFontSizeRecipient, "Text", nameof(AppSettings.RecipientFontsize));

        // --- Tab 3: Absender ---
        Bind(ckbPrintSender, "Checked", nameof(AppSettings.PrintSender));
        Bind(ckbBoldSender, "Checked", nameof(AppSettings.PrintSenderBold));
        Bind(tcSender, "SelectedIndex", nameof(AppSettings.SenderIndex));
        Bind(nudSenderOffsetX, "Value", nameof(AppSettings.SenderOffsetX));
        Bind(nudSenderOffsetY, "Value", nameof(AppSettings.SenderOffsetY));

        Bind(tbSender1, "Text", nameof(AppSettings.SenderLines1Joined));
        Bind(tbSender2, "Text", nameof(AppSettings.SenderLines2Joined));
        Bind(tbSender3, "Text", nameof(AppSettings.SenderLines3Joined));
        Bind(tbSender4, "Text", nameof(AppSettings.SenderLines4Joined));
        Bind(tbSender5, "Text", nameof(AppSettings.SenderLines5Joined));
        Bind(tbSender6, "Text", nameof(AppSettings.SenderLines6Joined));

        // --- Tab 4: Empfänger ---
        Bind(ckbBoldRecipient, "Checked", nameof(AppSettings.PrintRecipientBold));
        Bind(ckbAnredePrint, "Checked", nameof(AppSettings.PrintRecipientSalutation));
        Bind(ckbAnredeOberhalb, "Checked", nameof(AppSettings.RecipientSalutationAbove));
        Bind(ckbLandPrint, "Checked", nameof(AppSettings.PrintRecipientCountry));
        Bind(ckbLandGROSS, "Checked", nameof(AppSettings.RecipientCountryUpper));

        Bind(nudRecipOffsetX, "Value", nameof(AppSettings.RecipientOffsetX));
        Bind(nudRecipOffsetY, "Value", nameof(AppSettings.RecipientOffsetY));

        Bind(nudLineHeightFactor, "Value", nameof(AppSettings.LineHeightFactor));
        Bind(nudZipGapFactor, "Value", nameof(AppSettings.ZipGapFactor));
        Bind(nudLandGapFactor, "Value", nameof(AppSettings.LandGapFactor));
    }

    private void InitializePrinterSelection()
    {
        if (!string.IsNullOrEmpty(_settings.PrintDevice) && Utils.IsPrinterAvailable(_settings.PrintDevice))
        {
            printDocument.PrinterSettings.PrinterName = _settings.PrintDevice;
            if (printDocument.PrinterSettings.IsValid)
            {
                // Papierschächte laden
                cbSources.Items.Clear();
                foreach (PaperSource ps in printDocument.PrinterSettings.PaperSources) { cbSources.Items.Add(ps.SourceName); }

                // Papierformate laden
                cbPapersize.Items.Clear();
                foreach (PaperSize ps in printDocument.PrinterSettings.PaperSizes) { if (ps.Kind != PaperKind.Custom) { cbPapersize.Items.Add(ps.PaperName); } }

                // Initial Landscape setzen für Preview
                printDocument.DefaultPageSettings.Landscape = _settings.PrintLandscape;

                // Papierformat im PrintDocument setzen (wichtig für Preview)
                if (!string.IsNullOrEmpty(_settings.PrintFormat))
                {
                    foreach (PaperSize size in printDocument.PrinterSettings.PaperSizes)
                    {
                        if (size.PaperName == _settings.PrintFormat)
                        {
                            printDocument.DefaultPageSettings.PaperSize = size;
                            break;
                        }
                    }
                }
            }
        }
        else
        {
            if (string.IsNullOrEmpty(_settings.PrintDevice)) { Utils.MsgTaskDlg(Handle, "Druckerfehler", $"Es wurde noch kein Drucker ausgewählt."); }
            else { Utils.MsgTaskDlg(Handle, "Druckerfehler", $"Der Drucker '{_settings.PrintDevice}' ist nicht verfügbar."); }
        }
    }

    private void UpdateUiState()
    {
        // Logik für Abhängigkeiten
        ckbAnredeOberhalb.Enabled = ckbAnredePrint.Checked;

        var landActive = ckbLandPrint.Checked;
        lblLandGapFactor.Enabled = landActive;
        nudLandGapFactor.Enabled = landActive;
        lblLandRows.Enabled = landActive;
        ckbLandGROSS.Enabled = landActive;

        // Portrait/Landscape Visualisierung
        picPortrait.BorderStyle = rbPortrait.Checked ? BorderStyle.FixedSingle : BorderStyle.None;
        picLandscape.BorderStyle = rbLandscape.Checked ? BorderStyle.FixedSingle : BorderStyle.None;
    }

    // --- Event Handler ---
    private void PrintPreviewControl_ZoomChanged(object sender, EventArgs e) => UpdateZoomDisplay();

    private void BtnSave_Click(object sender, EventArgs e) => printDocument.Print();

    // Zentraler Handler für UI-Änderungen, die das Preview neu zeichnen müssen
    private void GenericControl_ValueChanged(object sender, EventArgs e)
    {
        // Durch DataBinding ist _settings bereits aktuell, wenn DataSourceUpdateMode.OnPropertyChanged genutzt wurde.
        // Wir müssen nur UI-Status updaten und Preview neu generieren.

        if (sender == rbLandscape || sender == rbPortrait)
        {
            // Spezielle Behandlung für RadioButtons, da Binding manchmal tricky ist bei Gruppen
            _settings.PrintLandscape = rbLandscape.Checked;
            printDocument.DefaultPageSettings.Landscape = rbLandscape.Checked;
        }

        UpdateUiState();
        printPreviewControl.GeneratePreviewSilently();
    }

    // Spezifische Handler, die mehr Logik brauchen als nur "Repaint"

    private void CbPrinter_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!cbPrinter.Visible || !cbPrinter.Focused) { return; }

        // Binding aktualisiert _settings automatisch, aber wir müssen printDocument validieren
        printDocument.PrinterSettings.PrinterName = cbPrinter.Text;

        if (!printDocument.PrinterSettings.IsValid)
        {
            Utils.MsgTaskDlg(Handle, "Druckerfehler", $"Der Drucker '{cbPrinter.Text}' ist nicht gültig.");
            return;
        }

        // Papierformate neu laden für gewählten Drucker
        cbPapersize.Items.Clear();
        foreach (PaperSize ps in printDocument.PrinterSettings.PaperSizes)
        {
            if (ps.Kind != PaperKind.Custom) { cbPapersize.Items.Add(ps.PaperName); }
        }

        // Papierschächte neu laden
        cbSources.Items.Clear();
        foreach (PaperSource ps in printDocument.PrinterSettings.PaperSources) { cbSources.Items.Add(ps.SourceName); }

        // Defaults setzen falls leer
        if (cbPapersize.Items.Count > 0) { cbPapersize.SelectedIndex = 0; }
        if (cbSources.Items.Count > 0) { cbSources.SelectedIndex = 0; }

        printPreviewControl.GeneratePreviewSilently();
        UpdateStatusBarInfo();
    }

    private void CbSources_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!cbSources.Visible || !cbSources.Focused) { return; }

        // Mapping Name -> PaperSource Objekt
        if (cbSources.SelectedIndex >= 0 && cbSources.SelectedIndex < printDocument.PrinterSettings.PaperSources.Count)
        {
            printDocument.DefaultPageSettings.PaperSource = printDocument.PrinterSettings.PaperSources[cbSources.SelectedIndex];
        }
        printPreviewControl.GeneratePreviewSilently();
    }

    private void CbPapersize_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!cbPapersize.Visible || !cbPapersize.Focused) { return; }

        // Mapping Name -> PaperSize Objekt
        if (cbPapersize.SelectedIndex >= 0 && cbPapersize.SelectedIndex < printDocument.PrinterSettings.PaperSizes.Count)
        {
            printDocument.DefaultPageSettings.PaperSize = printDocument.PrinterSettings.PaperSizes[cbPapersize.SelectedIndex];
        }
        printPreviewControl.GeneratePreviewSilently();
        UpdateStatusBarInfo();
    }

    private void RbPortrait_CheckedChanged(object sender, EventArgs e)
    {
        // Wird durch GenericControl_ValueChanged mitbehandelt, aber hier explizit für UI Logik
        printDocument.DefaultPageSettings.Landscape = !rbPortrait.Checked;
        UpdateUiState();
        printPreviewControl.GeneratePreviewSilently();
    }

    private void PicPortrait_Click(object sender, EventArgs e) => rbPortrait.Checked = true;
    private void PicLandscape_Click(object sender, EventArgs e) => rbLandscape.Checked = true;

    // --- Standard Funktionalität (Laden, Zeichnen, Tasten) ---

    private void FrmPrintSetting_Load(object sender, EventArgs e)
    {
        tbSender1.SetInnerMargins(8, 8);
        tbSender2.SetInnerMargins(8, 8);
        tbSender3.SetInnerMargins(8, 8);
        tbSender4.SetInnerMargins(8, 8);
        tbSender5.SetInnerMargins(8, 8);
        tbSender6.SetInnerMargins(8, 8);
        printPreviewControl.Document = printDocument;
        printPreviewControl.Zoom = _zoom;
        UpdateStatusBarInfo();
        UpdateZoomDisplay();
    }

    private void UpdateStatusBarInfo()
    {
        if (printDocument.PrinterSettings.IsValid && printDocument.DefaultPageSettings.PaperSize != null)
        {
            toolStripStatusLabel.Text = $"Format: {printDocument.DefaultPageSettings.PaperSize.PaperName}";
        }
        else { toolStripStatusLabel.Text = "Kein gültiger Drucker ausgewählt."; }
        UpdateStatusBarLayout();
    }

    private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
    {
        var g = e.Graphics;
        if (g == null) { return; }
        g.PageUnit = GraphicsUnit.Display;
        float lineH;
        var fontName = "Arial"; // Fallback
        if (!string.IsNullOrEmpty(_settings.PrintFont))
        {
            var parts = _settings.PrintFont.Split(',');
            fontName = parts[0].Trim();
        }

        // --- Absender Block ---
        if (_settings.PrintSender)
        {
            var style = _settings.PrintSenderBold ? FontStyle.Bold : FontStyle.Regular;
            using var fntSender = new Font(fontName, _settings.SenderFontsize, style);
            var senderXPos = e.MarginBounds.Left + (float)_settings.SenderOffsetX;
            var senderYPos = e.MarginBounds.Top + (float)_settings.SenderOffsetY;
            lineH = fntSender.GetHeight(g);
            var linesToPrint = _settings.SenderIndex switch
            {
                0 => _settings.SenderLines1,
                1 => _settings.SenderLines2,
                2 => _settings.SenderLines3,
                3 => _settings.SenderLines4,
                4 => _settings.SenderLines5,
                5 => _settings.SenderLines6,
                _ => _settings.SenderLines1
            };
            if (linesToPrint != null)
            {
                foreach (var line in linesToPrint)
                {
                    g.DrawString(line, fntSender, Brushes.Black, senderXPos, senderYPos);
                    senderYPos += lineH;
                }
            }
        }

        // --- Empfänger Block ---
        var recipStyle = _settings.PrintRecipientBold ? FontStyle.Bold : FontStyle.Regular;
        using var fntRecipient = new Font(fontName, _settings.RecipientFontsize, recipStyle);
        using var sf = new StringFormat();
        var lineHeightFactor = (float)_settings.LineHeightFactor;
        lineH = (float)(fntRecipient.GetHeight(g) * lineHeightFactor);
        var recipXPos = e.MarginBounds.Left + (e.MarginBounds.Width / 2) + (float)_settings.RecipientOffsetX;
        var recipYPos = e.MarginBounds.Top + (e.MarginBounds.Height / 2) + (float)_settings.RecipientOffsetY;
        var recipientLines = new string[6];
        var line1 = string.Empty;
        if (_settings.PrintRecipientSalutation && _recipientDict.TryGetValue("Anrede", out var anrede) && !string.IsNullOrEmpty(anrede)) { line1 += anrede; }
        if (_recipientDict.TryGetValue("Titel", out var titel) && !string.IsNullOrEmpty(titel)) { line1 += (line1.Length > 0 ? " " : "") + titel; }
        recipientLines[0] = line1.Trim();
        var line2 = string.Empty;
        if (_recipientDict.TryGetValue("Praefix", out var praefix) && !string.IsNullOrEmpty(praefix)) { line2 += praefix + " "; }
        if (_recipientDict.TryGetValue("Vorname", out var vorname) && !string.IsNullOrEmpty(vorname)) { line2 += vorname + " "; }
        if (_recipientDict.TryGetValue("Nachname", out var nachname) && !string.IsNullOrEmpty(nachname)) { line2 += nachname; }
        recipientLines[1] = line2.Trim();
        recipientLines[2] = _recipientDict.TryGetValue("Firma", out var firma) && !string.IsNullOrEmpty(firma) ? firma : string.Empty;
        recipientLines[3] = _recipientDict.TryGetValue("Strasse", out var strasse) && !string.IsNullOrEmpty(strasse) ? strasse : string.Empty;
        var line5 = string.Empty;
        if (_recipientDict.TryGetValue("PLZ", out var plz) && !string.IsNullOrEmpty(plz)) { line5 += plz + " "; }
        if (_recipientDict.TryGetValue("Ort", out var ort) && !string.IsNullOrEmpty(ort)) { line5 += ort; }
        recipientLines[4] = line5.Trim();
        if (_recipientDict.TryGetValue("Land", out var land) && _settings.PrintRecipientCountry && !string.IsNullOrEmpty(land)) { recipientLines[5] = _settings.RecipientCountryUpper ? land.ToUpper() : land; }
        else { recipientLines[5] = string.Empty; }
        if (!string.IsNullOrWhiteSpace(recipientLines[0]) && _settings.RecipientSalutationAbove) { recipYPos -= lineH; } // Korrektur Y-Position, wenn Anrede oberhalb gewünscht
        for (var i = 0; i < recipientLines.Length; i++) // Druck-Schleife Empfänger
        {
            if (string.IsNullOrWhiteSpace(recipientLines[i])) { continue; }
            if (i == 4) // Vor PLZ/Ort (Index 4) -> Abstand erhöhen
            {
                var zipGap = (float)_settings.ZipGapFactor;
                recipYPos += lineH * zipGap;
            }
            if (i == 5) // Vor Land (Index 5) -> Abstand erhöhen
            {
                var landGap = (float)_settings.LandGapFactor;
                recipYPos += lineH * landGap;
            }
            g.DrawString(recipientLines[i], fntRecipient, Brushes.Black, recipXPos, recipYPos, sf);
            recipYPos += lineH;
        }
        e.HasMorePages = false;
    }

    private void PrintDocument_BeginPrint(object sender, PrintEventArgs e) => printDocument.DocumentName = "Briefumschlag";

    private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tabControl.SelectedIndex == 0 && printDocument.DefaultPageSettings.PaperSize != null) { toolStripStatusLabel.Text = $"Format: {printDocument.DefaultPageSettings.PaperSize.PaperName}"; }
        else if (tabControl.SelectedIndex == 1) { toolStripStatusLabel.Text = $"Drucker: {printDocument.PrinterSettings.PrinterName}"; }
        else if (printDocument.DefaultPageSettings.PaperSize != null) { toolStripStatusLabel.Text = $"{printDocument.PrinterSettings.PrinterName}, {printDocument.DefaultPageSettings.PaperSize.PaperName}"; }
    }

    private void UpdateStatusBarLayout()
    {
        if (lblZoomStatus != null && printPreviewControl != null) { lblZoomStatus.Width = printPreviewControl.Width; }
    }

    private void UpdateZoomDisplay() => lblZoomStatus?.Text = $"Zoom: {printPreviewControl.Zoom * 100:0}% (Doppelklick, Mausrad oder Strg++/-/0 für Änderung)";  // printPreviewControl_ZoomChanged
    private void FrmPrintSetting_Layout(object sender, LayoutEventArgs e) => UpdateStatusBarLayout();

    private void ZoomInToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (printPreviewControl.Zoom < 1D) { printPreviewControl.Zoom += 0.1; if (printPreviewControl.Zoom > 1) { printPreviewControl.Zoom = 1D; } }
    }

    private void ZoomOutToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (printPreviewControl.Zoom >= 0.3D) { printPreviewControl.Zoom -= 0.1; }
    }

    private void ZoomDefaultToolStripMenuItem_Click(object sender, EventArgs e) => printPreviewControl.Zoom = _zoom;

    private void ContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (printPreviewControl.Zoom == _zoom) { zoomDefaultToolStripMenuItem.Enabled = false; }
        else if (printPreviewControl.Zoom < 0.3D) { zoomOutToolStripMenuItem.Enabled = false; }
        else if (printPreviewControl.Zoom > 1D) { zoomInToolStripMenuItem.Enabled = false; }
        else
        {
            zoomDefaultToolStripMenuItem.Enabled = true;
            zoomInToolStripMenuItem.Enabled = true;
            zoomOutToolStripMenuItem.Enabled = true;
        }
    }

    private void TcSender_DrawItem(object sender, DrawItemEventArgs e)
    {
        if (sender is TabControl tabControlSender)
        {
            using var g = e.Graphics;
            Brush textBrush;
            e.DrawBackground();
            var tabPage = tabControlSender.TabPages[e.Index];
            if (e.State == DrawItemState.Selected)
            {
                textBrush = new SolidBrush(Color.White);
                g.FillRectangle(Brushes.Gray, e.Bounds);
            }
            else
            {
                textBrush = new SolidBrush(e.ForeColor);
                e.DrawBackground();
            }
            using var tabFont = new Font("Segoe UI", 10.0f, FontStyle.Bold, GraphicsUnit.Point);
            using var stringFlags = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            var tabBounds = tabControlSender.GetTabRect(e.Index);
            tabBounds.Inflate(-2, -2);
            g.DrawString(tabPage.Text, tabFont, textBrush, tabBounds, new StringFormat(stringFlags));
        }
    }

    // TextBox Debounce (Wichtig für Performance beim Tippen)
    private void TbSender_TextChanged(object sender, EventArgs e)
    {
        timerDebounce.Stop();
        timerDebounce.Start();
    }

    private void TimerDebounce_Tick(object sender, EventArgs e)
    {
        timerDebounce.Stop();
        printPreviewControl.GeneratePreviewSilently();
    }

    private void PrintPreviewControl_DoubleClick(object? sender, EventArgs e)
    {
        var fitZoom = GetBestFitZoom();
        var currentZoom = printPreviewControl.Zoom;
        if (Math.Abs(currentZoom - 1.0) < Math.Abs(currentZoom - fitZoom)) { printPreviewControl.Zoom = fitZoom; }
        else { printPreviewControl.Zoom = 1.0; }
        UpdateZoomDisplay();
    }

    private double GetBestFitZoom()
    {
        if (printDocument.DefaultPageSettings.PaperSize == null) { return 1.0; }
        double clientWidth = printPreviewControl.ClientSize.Width - 25;
        double clientHeight = printPreviewControl.ClientSize.Height - 25;
        if (clientWidth <= 0 || clientHeight <= 0) { return 0.1; }
        var paperSize = printDocument.DefaultPageSettings.PaperSize;
        var isLandscape = printDocument.DefaultPageSettings.Landscape;
        double paperW = paperSize.Width;
        double paperH = paperSize.Height;
        if (isLandscape) { (paperW, paperH) = (paperH, paperW); }
        var paperPixelWidth = paperW / 100.0 * 96.0;
        var paperPixelHeight = paperH / 100.0 * 96.0;
        var zoomX = clientWidth / paperPixelWidth;
        var zoomY = clientHeight / paperPixelHeight;
        return Math.Max(Math.Min(zoomX, zoomY), 0.1);
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        switch (keyData)
        {
            case Keys.Tab:
                tabControl.SelectedIndex = (tabControl.SelectedIndex + 1) % tabControl.TabCount;
                return true;
            case Keys.Oemplus | Keys.Control:
            case Keys.Add | Keys.Control:
                if (printPreviewControl.Zoom < 1.0)
                {
                    printPreviewControl.Zoom = Math.Min(1.0, printPreviewControl.Zoom + 0.1);
                    UpdateZoomDisplay();
                }
                else { Console.Beep(); }
                return true;
            case Keys.OemMinus | Keys.Control:
            case Keys.Subtract | Keys.Control:
                if (printPreviewControl.Zoom > 0.3)
                {
                    printPreviewControl.Zoom = Math.Max(0.3, printPreviewControl.Zoom - 0.1);
                    UpdateZoomDisplay();
                }
                else { Console.Beep(); }
                return true;
            case Keys.NumPad0 | Keys.Control:
            case Keys.D0 | Keys.Control:
                var bestFit = GetBestFitZoom();
                if (Math.Abs(printPreviewControl.Zoom - bestFit) < 0.001) { Console.Beep(); }
                else { printPreviewControl.Zoom = bestFit; UpdateZoomDisplay(); }
                return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

}