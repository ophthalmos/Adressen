using System.Drawing.Printing;
using Adressen.cls;

namespace Adressen;

public partial class FrmPrintSetting : Form
{
    public string Device => cbPrinter.Text; // printDocument.PrinterSettings.PrinterName;
    public string Source => cbSources.Text;
    public bool Landscape => rbLandscape.Checked;
    public string Format => cbPapersize.Text;
    public string Schrift => cbFont.Text;
    public int SenderSize => int.Parse(cbFontsizeSender.Text);
    public int RecipSize => int.Parse(cbFontSizeRecipient.Text);
    public int SenderIndex => tcSender.SelectedIndex;  // nullbasiert
    public string[] SenderLines1 => tbSender1.Lines;
    public string[] SenderLines2 => tbSender2.Lines;
    public string[] SenderLines3 => tbSender3.Lines;
    public string[] SenderLines4 => tbSender4.Lines;
    public string[] SenderLines5 => tbSender5.Lines;
    public string[] SenderLines6 => tbSender6.Lines;
    public bool SenderPrint => ckbPrintSender.Checked;
    public decimal RecipX => nudRecipOffsetX.Value;
    public decimal RecipY => nudRecipOffsetY.Value;
    public decimal SendX => nudSenderOffsetX.Value;
    public decimal SendY => nudSenderOffsetY.Value;
    public bool RecipBold => ckbBoldRecipient.Checked;
    public bool SendBold => ckbBoldSender.Checked;
    public bool Salutation => ckbAnredePrint.Checked;
    public bool SalutAbove => ckbAnredeOberhalb.Checked;
    public bool Country => ckbLandPrint.Checked;
    public bool CountryUpper => ckbLandGROSS.Checked;
    public decimal LineHeightFactor => nudLineHeightFactor.Value;
    public decimal ZipGapFactor => nudZipGapFactor.Value;
    public decimal LandGapFactor => nudLandGapFactor.Value;

    private readonly Dictionary<string, string> _recipientDict;
    private readonly double _zoom = 0.50F; // double ist korrekt

    public FrmPrintSetting(string colorScheme, Dictionary<string, string> recipientDict,
                           string pDevice, string pSource, bool pLandscape,
                           string pFormat, string pFont, int pSenderSize, int pRecipSize,
                           int pSenderIndex, string[] pSenderLines1, string[] pSenderLines2, string[] pSenderLines3, string[] pSenderLines4, string[] pSenderLines5, string[] pSenderLines6, bool pSenderPrint,
                           decimal pRecipX, decimal pRecipY, decimal pSendX, decimal pSendY, bool pRecipBold, bool pSendBold, bool pAnrede, bool pAbove, bool pLand, bool pUpper, decimal pLineHeight, decimal pZipGap, decimal pLandGap)
    {
        InitializeComponent();
        printPreviewControl.ZoomChanged += (s, e) => UpdateZoomDisplay();
        foreach (TabPage tabPage in tabControl.TabPages)
        {
            tabPage.BackColor = colorScheme switch
            {
                "blue" => SystemColors.InactiveBorder,
                "pale" => SystemColors.ControlLightLight,
                "dark" => SystemColors.AppWorkspace,
                _ => SystemColors.ButtonFace,
            };
        }
        _recipientDict = recipientDict;
        cbFont.Items.AddRange([.. FontFamily.Families.Select(f => f.Name)]);
        foreach (string s in PrinterSettings.InstalledPrinters) { cbPrinter.Items.Add(s); }
        tcSender.SelectedIndex = pSenderIndex;
        tbSender1.Lines = pSenderLines1;
        tbSender2.Lines = pSenderLines2;
        tbSender3.Lines = pSenderLines3;
        tbSender4.Lines = pSenderLines4;
        tbSender5.Lines = pSenderLines5;
        tbSender6.Lines = pSenderLines6;
        printDocument.DefaultPageSettings.Margins = new Margins(25, 25, 25, 25); // hundredths of an inch, default 100 (2,54 cm für alle Seitenränder)
        cbFont.Text = pFont; // "Calibri";
        cbFontSizeRecipient.Text = pRecipSize.ToString(); // "14";
        cbFontsizeSender.Text = pSenderSize.ToString(); // "12";

        ckbPrintSender.Checked = pSenderPrint;
        nudRecipOffsetX.Value = pRecipX;
        nudRecipOffsetY.Value = pRecipY;
        nudSenderOffsetX.Value = pSendX;
        nudSenderOffsetY.Value = pSendY;
        ckbBoldRecipient.Checked = pRecipBold;
        ckbBoldSender.Checked = pSendBold;
        ckbAnredePrint.Checked = pAnrede;
        ckbAnredeOberhalb.Checked = pAbove;
        ckbLandPrint.Checked = pLand;
        ckbLandGROSS.Checked = pUpper;
        nudLineHeightFactor.Value = pLineHeight;
        nudZipGapFactor.Value = pZipGap;
        nudLandGapFactor.Value = pLandGap;

        if (ckbAnredePrint.Checked) { ckbAnredeOberhalb.Enabled = true; }
        else { ckbAnredeOberhalb.Enabled = false; }
        if (ckbLandPrint.Checked) { lblLandGapFactor.Enabled = nudLandGapFactor.Enabled = lblLandRows.Enabled = ckbLandGROSS.Enabled = true; }
        else { lblLandGapFactor.Enabled = nudLandGapFactor.Enabled = lblLandRows.Enabled = ckbLandGROSS.Enabled = false; }

        if (!string.IsNullOrEmpty(pDevice) && Utils.IsPrinterAvailable(pDevice))
        {
            printDocument.PrinterSettings.PrinterName = pDevice;
            if (printDocument.PrinterSettings.IsValid)
            {
                cbPrinter.Text = printDocument.PrinterSettings.PrinterName;
                foreach (PaperSource ps in printDocument.PrinterSettings.PaperSources) { cbSources.Items.Add(ps.SourceName); }
                cbSources.Text = pSource; // "Papiereinzug hinten";
                printDocument.DefaultPageSettings.Landscape = pLandscape;
                if (printDocument.DefaultPageSettings.Landscape) { rbLandscape.Checked = true; }
                else { rbPortrait.Checked = true; }
                foreach (PaperSize ps in printDocument.PrinterSettings.PaperSizes)
                {
                    if (ps.Kind != PaperKind.Custom) { cbPapersize.Items.Add(ps.PaperName); } // Nur Standardformate anzeigen
                }

                if (printDocument.DefaultPageSettings.PaperSize != null)
                {

                    cbPapersize.Text = printDocument.DefaultPageSettings.PaperSize.PaperName; // wird ggfs. später überschrieben    
                    foreach (PaperSize size in printDocument.PrinterSettings.PaperSizes)
                    {
                        if (size.PaperName == pFormat)
                        {
                            printDocument.DefaultPageSettings.PaperSize = size;
                            cbPapersize.Text = pFormat;
                            break;
                        }
                    }
                }
            }
        }
        else
        {
            if (string.IsNullOrEmpty(pDevice)) { Utils.MsgTaskDlg(Handle, "Druckerfehler", $"Es wurde noch kein Drucker ausgewählt."); }
            else { Utils.MsgTaskDlg(Handle, "Druckerfehler", $"Der Drucker '{pDevice}' ist nicht verfügbar."); }
            return;
        }

    }

    private void BtnSave_Click(object sender, EventArgs e) => printDocument.Print();

    private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
    {
        var g = e.Graphics;
        if (g == null) { return; }
        g.PageUnit = GraphicsUnit.Display;
        float lineH;
        // --- Absender Block ---
        if (ckbPrintSender.Checked)
        {
            using var fntSender = new Font(cbFont.Text, float.Parse(cbFontsizeSender.Text), ckbBoldSender.Checked ? FontStyle.Bold : FontStyle.Regular);
            var senderXPos = e.MarginBounds.Left + (float)nudSenderOffsetX.Value;
            var senderYPos = e.MarginBounds.Top + (float)nudSenderOffsetY.Value;
            lineH = fntSender.GetHeight(g);
            var controls = tcSender.TabPages[tcSender.SelectedIndex].Controls;
            if (controls.Count == 1 && controls[0] is TextBox tb)
            {
                for (var i = 0; i < tb.Lines.Length; i++)
                {
                    g.DrawString(tb.Lines[i], fntSender, Brushes.Black, senderXPos, senderYPos);
                    senderYPos += lineH;
                }
            }
        }
        // --- Empfänger Block ---
        using var fntRecipient = new Font(cbFont.Text, float.Parse(cbFontSizeRecipient.Text), ckbBoldRecipient.Checked ? FontStyle.Bold : FontStyle.Regular);
        using var sf = new StringFormat();
        var lineHeightFactor = (float)nudLineHeightFactor.Value; // variabler Zeilenabstand
        lineH = (float)(fntRecipient.GetHeight(g) * lineHeightFactor);

        var recipXPos = e.MarginBounds.Left + (e.MarginBounds.Width / 2) + (float)nudRecipOffsetX.Value;
        var recipYPos = e.MarginBounds.Top + (e.MarginBounds.Height / 2) + (float)nudRecipOffsetY.Value;

        var recipientLines = new string[6];
        recipientLines[0] = (_recipientDict.TryGetValue("Anrede", out var anrede) && ckbAnredePrint.Checked && !string.IsNullOrEmpty(anrede) ? anrede : string.Empty)
                          + (_recipientDict.TryGetValue("Titel", out var titel) && !string.IsNullOrEmpty(titel) ? " " + titel : string.Empty);
        recipientLines[0] = recipientLines[0].Trim();
        recipientLines[1] = (_recipientDict.TryGetValue("Praefix", out var praefix) && !string.IsNullOrEmpty(praefix) ? praefix + " " : string.Empty)
                          + (_recipientDict.TryGetValue("Vorname", out var vorname) && !string.IsNullOrEmpty(vorname) ? vorname + " " : string.Empty)
                          + (_recipientDict.TryGetValue("Nachname", out var nachname) && !string.IsNullOrEmpty(nachname) ? nachname : string.Empty);
        recipientLines[2] = _recipientDict.TryGetValue("Firma", out var firma) && !string.IsNullOrEmpty(firma) ? firma : string.Empty;
        recipientLines[3] = _recipientDict.TryGetValue("Strasse", out var strasse) && !string.IsNullOrEmpty(strasse) ? strasse : string.Empty;
        recipientLines[4] = (_recipientDict.TryGetValue("PLZ", out var plz) && !string.IsNullOrEmpty(plz) ? plz + " " : string.Empty)
                          + (_recipientDict.TryGetValue("Ort", out var ort) && !string.IsNullOrEmpty(ort) ? ort : string.Empty);
        if (_recipientDict.TryGetValue("Land", out var land) && ckbLandPrint.Checked && !string.IsNullOrEmpty(land)) { recipientLines[5] = ckbLandGROSS.Checked ? land.ToUpper() : land; }
        else { recipientLines[5] = string.Empty; }
        if (!string.IsNullOrWhiteSpace(recipientLines[0]) && ckbAnredeOberhalb.Checked) { recipYPos -= lineH; } // Platz für Anrede oberhalb schaffen
        for (var i = 0; i < recipientLines.Length; i++)  // Zeichnen der Empfängerzeilen
        {
            if (string.IsNullOrWhiteSpace(recipientLines[i])) { continue; }
            if (i == 4) // PLZ-Ort-Zeile
            {
                var zipGap = (float)nudZipGapFactor.Value;
                recipYPos += lineH * zipGap;
            }
            if (i == 5) // Land-Zeile
            {
                var landGap = (float)nudLandGapFactor.Value;
                recipYPos += lineH * landGap;
            }
            g.DrawString(recipientLines[i], fntRecipient, Brushes.Black, recipXPos, recipYPos, sf);
            recipYPos += lineH;  // Nächste Zeile berechnen
        }
        e.HasMorePages = false;
    }

    private void CbPrinter_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!cbPrinter.Visible || !cbPrinter.Focused) { return; }

        printDocument.PrinterSettings.PrinterName = cbPrinter.Text;
        if (!printDocument.PrinterSettings.IsValid)
        {
            Utils.MsgTaskDlg(Handle, "Druckerfehler", $"Der Drucker '{cbPrinter.Text}' ist nicht gültig.");
            return;
        }

        printDocument.DefaultPageSettings.PaperSize = printDocument.PrinterSettings.DefaultPageSettings.PaperSize;
        if (printDocument.DefaultPageSettings.PaperSize != null)
        {
            cbPapersize.Items.Clear();
            foreach (PaperSize ps in printDocument.PrinterSettings.PaperSizes)
            {
                if (ps.Kind != PaperKind.Custom) { cbPapersize.Items.Add(ps.PaperName); }
            }
            cbPapersize.SelectedItem = printDocument.PrinterSettings.DefaultPageSettings.PaperSize.PaperName;
        }
        else { cbPapersize.SelectedIndex = 0; }

        cbSources.Items.Clear();
        foreach (PaperSource ps in printDocument.PrinterSettings.PaperSources) { cbSources.Items.Add(ps.SourceName); }
        printDocument.DefaultPageSettings.PaperSource = printDocument.PrinterSettings.DefaultPageSettings.PaperSource;
        if (printDocument.DefaultPageSettings.PaperSource != null)
        {
            cbSources.SelectedItem = printDocument.DefaultPageSettings.PaperSource.SourceName;
        }
        else { cbSources.SelectedIndex = 0; }
        printPreviewControl.GeneratePreviewSilently();
        if (printDocument.DefaultPageSettings.PaperSize != null) { toolStripStatusLabel.Text = $"Format: {printDocument.DefaultPageSettings.PaperSize.PaperName}"; }
    }

    private void CbSources_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!cbSources.Visible || !cbSources.Focused) { return; } // Keine Auswahl getroffen 
        printDocument.DefaultPageSettings.PaperSource = printDocument.PrinterSettings.PaperSources[cbSources.SelectedIndex] ?? printDocument.DefaultPageSettings.PaperSource;
        printPreviewControl.GeneratePreviewSilently();
        printPreviewControl.Zoom = _zoom; // 48 / 100f;
    }

    private void RbPortrait_CheckedChanged(object sender, EventArgs e)
    {
        printDocument.DefaultPageSettings.Landscape = !rbPortrait.Checked;
        picPortrait.BorderStyle = rbPortrait.Checked ? BorderStyle.FixedSingle : BorderStyle.None;
        picLandscape.BorderStyle = rbPortrait.Checked ? BorderStyle.None : BorderStyle.FixedSingle;
        printPreviewControl.GeneratePreviewSilently();
        printPreviewControl.Zoom = _zoom; // 48 / 100f;
    }

    private void CbPapersize_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!cbPapersize.Visible || !cbPapersize.Focused) { return; } // Keine Auswahl getroffen 

        if (cbPapersize.SelectedIndex >= 0 && cbPapersize.SelectedIndex < printDocument.PrinterSettings.PaperSizes.Count)
        {
            printDocument.DefaultPageSettings.PaperSize = printDocument.PrinterSettings.PaperSizes[cbPapersize.SelectedIndex];
        }
        printPreviewControl.GeneratePreviewSilently();
        printPreviewControl.Zoom = _zoom; // 48 / 100f;
    }

    private void PrintDocument_BeginPrint(object sender, PrintEventArgs e) => printDocument.DocumentName = "Briefumschlag";

    private void FrmPrintSetting_Load(object sender, EventArgs e)
    {
        NativeMethods.SendMessage(tbSender1.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 8 << 16);
        NativeMethods.SendMessage(tbSender1.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 8);
        NativeMethods.SendMessage(tbSender2.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 8 << 16);
        NativeMethods.SendMessage(tbSender2.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 8);
        NativeMethods.SendMessage(tbSender3.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 8 << 16);
        NativeMethods.SendMessage(tbSender3.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 8);
        NativeMethods.SendMessage(tbSender4.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 8 << 16);
        NativeMethods.SendMessage(tbSender4.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 8);
        NativeMethods.SendMessage(tbSender5.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 8 << 16);
        NativeMethods.SendMessage(tbSender5.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 8);
        NativeMethods.SendMessage(tbSender6.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 8 << 16);
        NativeMethods.SendMessage(tbSender6.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 8);
        printPreviewControl.Document = printDocument;
        printPreviewControl.Zoom = _zoom; // 48 / 100f;
        if (printDocument.PrinterSettings.IsValid && printDocument.DefaultPageSettings.PaperSize != null)
        {
            toolStripStatusLabel.Text = $"Format: {printDocument.DefaultPageSettings.PaperSize.PaperName}";
        }
        else { toolStripStatusLabel.Text = "Kein gültiger Drucker ausgewählt."; }
        UpdateStatusBarLayout();
        UpdateZoomDisplay();
    }

    private void PicPortrait_Click(object sender, EventArgs e) => rbPortrait.Checked = true;
    private void PicLandscape_Click(object sender, EventArgs e) => rbLandscape.Checked = true;

    private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tabControl.SelectedIndex == 0) { toolStripStatusLabel.Text = $"Format: {printDocument.DefaultPageSettings.PaperSize.PaperName}"; }
        else if (tabControl.SelectedIndex == 1) { toolStripStatusLabel.Text = $"Drucker: {printDocument.PrinterSettings.PrinterName}"; }
        else { toolStripStatusLabel.Text = $"{printDocument.PrinterSettings.PrinterName}, {printDocument.DefaultPageSettings.PaperSize.PaperName}"; }
    }

    private void ZoomInToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (printPreviewControl.Zoom < 1D) { printPreviewControl.Zoom += 0.1; if (printPreviewControl.Zoom > 1) { printPreviewControl.Zoom = 1D; } }
    }

    private void ZoomOutToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (printPreviewControl.Zoom >= 0.3D) { printPreviewControl.Zoom -= 0.1; }
    }

    private void ZoomDefaultToolStripMenuItem_Click(object sender, EventArgs e)
    {
        printPreviewControl.Zoom = _zoom; // 48 / 100f;
    }

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
            var tabPage = tabControlSender.TabPages[e.Index]; // Get the item from the collection.
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
            using var stringFlags = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center }; // Draw string. Center the text.
            var tabBounds = tabControlSender.GetTabRect(e.Index); // Get the real bounds for the tab rectangle.
            tabBounds.Inflate(-2, -2); // Inflate the rectangle to make it smaller.
            g.DrawString(tabPage.Text, tabFont, textBrush, tabBounds, new StringFormat(stringFlags));
        }
    }

    private void GenericControl_ValueChanged(object sender, EventArgs e)
    {
        if (sender is Control c && c.Visible)
        {
            if (c == ckbAnredePrint && ckbAnredePrint.Checked) { ckbAnredeOberhalb.Enabled = true; }
            else if (c == ckbAnredePrint && !ckbAnredePrint.Checked) { ckbAnredeOberhalb.Enabled = false; }
            else if (c == ckbLandPrint && ckbLandPrint.Checked) { lblLandGapFactor.Enabled = nudLandGapFactor.Enabled = lblLandRows.Enabled = ckbLandGROSS.Enabled = true; }
            else if (c == ckbLandPrint && !ckbLandPrint.Checked) { lblLandGapFactor.Enabled = nudLandGapFactor.Enabled = lblLandRows.Enabled = ckbLandGROSS.Enabled = false; }
            printPreviewControl.GeneratePreviewSilently();
        }
    }

    private void TbSender_TextChanged(object sender, EventArgs e)
    {
        timerDebounce.Stop(); // Beim Tippen wollen wir warten, bis der Nutzer kurz Pause macht
        timerDebounce.Start();
    }

    private void TimerDebounce_Tick(object sender, EventArgs e)
    {
        timerDebounce.Stop();
        printPreviewControl.GeneratePreviewSilently();
    }

    private void UpdateStatusBarLayout()
    {
        if (lblZoomStatus != null && printPreviewControl != null) { lblZoomStatus.Width = printPreviewControl.Width; }
    }

    private void UpdateZoomDisplay() => lblZoomStatus?.Text = $"Zoom: {printPreviewControl.Zoom * 100:0}% (Doppelklick, Mausrad oder Strg++/-/0 für Änderung)";

    private void PrintPreviewControl_DoubleClick(object? sender, EventArgs e)
    {
        var fitZoom = GetBestFitZoom();
        var currentZoom = printPreviewControl.Zoom;
        if (Math.Abs(currentZoom - 1.0) < Math.Abs(currentZoom - fitZoom)) { printPreviewControl.Zoom = fitZoom; }
        else { printPreviewControl.Zoom = 1.0; } // wenn näher an 100%, Best Fit, sonst 100%
        UpdateZoomDisplay();
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
                else
                {
                    printPreviewControl.Zoom = bestFit;
                    UpdateZoomDisplay();
                }
                return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
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

    private void FrmPrintSetting_Layout(object sender, LayoutEventArgs e) => UpdateStatusBarLayout();  // Bei jedem Layout-Vorgang (Resize etc.) die Breite anpassen
}
