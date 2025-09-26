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
    public bool Country => ckbLandPrint.Checked;

    private readonly Dictionary<string, string> _recipientDict;
    private readonly double _zoom = 0.50F; // double ist korrekt

    public FrmPrintSetting(string colorScheme, Dictionary<string, string> recipientDict,
                           string pDevice, string pSource, bool pLandscape,
                           string pFormat, string pFont, int pSenderSize, int pRecipSize,
                           int pSenderIndex, string[] pSenderLines1, string[] pSenderLines2, string[] pSenderLines3, string[] pSenderLines4, string[] pSenderLines5, string[] pSenderLines6, bool pSenderPrint,
                           decimal pRecipX, decimal pRecipY, decimal pSendX, decimal pSendY, bool pRecipBold, bool pSendBold, bool pAnrede, bool pLand)
    {
        InitializeComponent();
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
        ckbLandPrint.Checked = pLand;

        if (!string.IsNullOrEmpty(pDevice) && Utilities.IsPrinterAvailable(pDevice))
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
            if (string.IsNullOrEmpty(pDevice)) { Utilities.ErrorMsgTaskDlg(Handle, "Druckerfehler", $"Es wurde noch kein Drucker ausgewählt."); }
            else { Utilities.ErrorMsgTaskDlg(Handle, "Druckerfehler", $"Der Drucker '{pDevice}' ist nicht verfügbar."); }
            return;
        }

    }

    private void BtnSave_Click(object sender, EventArgs e) => printDocument.Print();

    private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
    {
        using var g = e.Graphics;
        if (g == null) { return; }
        g.PageUnit = GraphicsUnit.Display; // 0.01 inch, 1/ 
        float lineH;
        if (ckbPrintSender.Checked)
        {
            using var fntSender = new Font(cbFont.Text, float.Parse(cbFontsizeSender.Text), (ckbBoldSender.Checked ? FontStyle.Bold : FontStyle.Regular)); //, GraphicsUnit.Point
            var senderXPos = e.MarginBounds.Left + (float)nudSenderOffsetX.Value; // 50
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
        using var fntRecipient = new Font(cbFont.Text, float.Parse(cbFontSizeRecipient.Text), ckbBoldRecipient.Checked ? FontStyle.Bold : FontStyle.Regular);
        using var sf = new StringFormat(); // StringFormatFlags.NoWrap
        //sf.LineAlignment = StringAlignment.Center; // Important for vertical centering
        lineH = (float)(fntRecipient.GetHeight(g) * 1.2);
        var recipXPos = e.MarginBounds.Left + (e.MarginBounds.Width / 2) + (float)nudRecipOffsetX.Value;
        var recipYPos = e.MarginBounds.Top + (e.MarginBounds.Height / 2) + (float)nudRecipOffsetY.Value;

        var _recipientLines = new string[6];
        _recipientLines[0] = (_recipientDict.TryGetValue("Anrede", out var value) && ckbAnredePrint.Checked && !string.IsNullOrEmpty(value) ? value : string.Empty)  // Anrede 
            + (_recipientDict.TryGetValue("Titel", out value) && !string.IsNullOrEmpty(value) ? value + " " : string.Empty); // Titel
        _recipientLines[1] = (_recipientDict.TryGetValue("Präfix", out value) && !string.IsNullOrEmpty(value) ? value + " " : string.Empty) // Präfix
            + (_recipientDict.TryGetValue("Vorname", out value) && !string.IsNullOrEmpty(value) ? value + " " : string.Empty) // Vorname
            + (_recipientDict.TryGetValue("Nachname", out value) && !string.IsNullOrEmpty(value) ? value : string.Empty); // Nachname
        _recipientLines[2] = _recipientDict.TryGetValue("Firma", out value) && !string.IsNullOrEmpty(value) ? value : string.Empty; // Firma    
        _recipientLines[3] = _recipientDict.TryGetValue("StraßeNr", out value) && !string.IsNullOrEmpty(value) ? value : string.Empty; // Straße  
        _recipientLines[4] = (_recipientDict.TryGetValue("PLZ", out value) && !string.IsNullOrEmpty(value) ? value + " " : string.Empty) // PLZ
            + (_recipientDict.TryGetValue("Ort", out value) && !string.IsNullOrEmpty(value) ? value : string.Empty); // Ort 
        _recipientLines[5] = _recipientDict.TryGetValue("Land", out value) && ckbLandPrint.Checked && !string.IsNullOrEmpty(value) ? value : string.Empty; // Land  

        for (var i = 0; i < _recipientLines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(_recipientLines[i])) { continue; } // Leere Zeilen überspringen   
            //if (i == 1 && string.IsNullOrWhiteSpace(_recipientLines[0])) { recipYPos -= lineH; }
            //else if (i == 2 && string.IsNullOrWhiteSpace(_recipientLines[1])) { recipYPos -= lineH; }
            else if (i == 4) { recipYPos += lineH / 3; } // Zeilenabstand erhöhen vor PLZ/Ort
            g.DrawString(_recipientLines[i], fntRecipient, Brushes.Black, recipXPos, recipYPos, sf);
            recipYPos += lineH;
        }

        e.HasMorePages = false;
    }

    private void CbPrinter_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!cbPrinter.Visible || !cbPrinter.Focused) { return; } // Keine Auswahl getroffen 
        printDocument.PrinterSettings.PrinterName = cbPrinter.Text;
        if (!printDocument.PrinterSettings.IsValid)
        {
            Utilities.ErrorMsgTaskDlg(Handle, "Druckerfehler", $"Der Drucker '{cbPrinter.Text}' ist nicht gültig oder nicht verfügbar.");
            return;
        }
        //MessageBox.Show(printDocument.PrinterSettings.DefaultPageSettings.PaperSize.PaperName + Environment.NewLine + printDocument.PrinterSettings.DefaultPageSettings.PaperSource.SourceName);

        printDocument.DefaultPageSettings.PaperSize = printDocument.PrinterSettings.DefaultPageSettings.PaperSize;
        if (printDocument.DefaultPageSettings.PaperSize != null)
        {
            cbPapersize.Items.Clear();
            foreach (PaperSize ps in printDocument.PrinterSettings.PaperSizes)
            {
                if (ps.Kind != PaperKind.Custom) { cbPapersize.Items.Add(ps.PaperName); } // Nur Standardformate anzeigen
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
            //MessageBox.Show($"PaperSource: {cbSources.SelectedItem}"); // Debug
            //     cbSources.Text = printDocument.DefaultPageSettings.PaperSource.SourceName; // Workaround, damit der Text in der ComboBox angezeigt wird
        }
        else
        {
            //printDocument.DefaultPageSettings.PaperSource = printDocument.PrinterSettings.PaperSources[0];
            cbSources.SelectedIndex = 0;
        }

        //var paperSizeCode = NativeMethods.GetDefaultPaperSize(printDocument.PrinterSettings.PrinterName);
        //var matchingSize = Utilities.GetMatchingPaperSize(printDocument, paperSizeCode);
        //if (paperSizeCode > 0 && matchingSize != null)
        //{
        //    printDocument.DefaultPageSettings.PaperSize = matchingSize;
        //    cbPapersize.SelectedItem = matchingSize.PaperName; // .Text = match.PaperName;
        //}
        //else
        //{
        //    printDocument.DefaultPageSettings.PaperSize = Utilities.GetDefaultPaperSize(printDocument.PrinterSettings, PaperKind.C6Envelope);
        //    cbPapersize.SelectedIndex = 0;
        //}
        //printPreviewControl.Document = null;
        printPreviewControl.Document = printDocument;
        printPreviewControl.Zoom = _zoom; // 48 / 100f;
        toolStripStatusLabel.Text = $"Format: {cbPapersize.Text}";
    }

    private void CbSources_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!cbSources.Visible || !cbSources.Focused) { return; } // Keine Auswahl getroffen 
        //if (!_lastPrinterName.Equals(printDocument.PrinterSettings.PrinterName))
        //{
        //    cbSources.Items.Clear();
        //    foreach (PaperSource ps in printDocument.PrinterSettings.PaperSources) { cbSources.Items.Add(ps.SourceName); }
        //    _lastPrinterName = printDocument.PrinterSettings.PrinterName;
        //    return; // Aktualisierung der Papiersourcen nur bei Druckerwechsel - Workaround weil es zu kompliziert ist, einen phsysikalischen Druckerwechsel zu erkennen    
        //}
        printDocument.DefaultPageSettings.PaperSource = printDocument.PrinterSettings.PaperSources[cbSources.SelectedIndex] ?? printDocument.DefaultPageSettings.PaperSource;
        printPreviewControl.Document = printDocument;
        printPreviewControl.Zoom = _zoom; // 48 / 100f;
    }

    private void RbPortrait_CheckedChanged(object sender, EventArgs e)
    {
        printDocument.DefaultPageSettings.Landscape = !rbPortrait.Checked;
        picPortrait.BorderStyle = rbPortrait.Checked ? BorderStyle.FixedSingle : BorderStyle.None;
        picLandscape.BorderStyle = rbPortrait.Checked ? BorderStyle.None : BorderStyle.FixedSingle;
        printPreviewControl.Document = printDocument;
        printPreviewControl.Zoom = _zoom; // 48 / 100f;
    }

    private void CbPapersize_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (!cbPapersize.Visible || !cbPapersize.Focused) { return; } // Keine Auswahl getroffen 

        if (cbPapersize.SelectedIndex >= 0 && cbPapersize.SelectedIndex < printDocument.PrinterSettings.PaperSizes.Count)
        {
            printDocument.DefaultPageSettings.PaperSize = printDocument.PrinterSettings.PaperSizes[cbPapersize.SelectedIndex];
        }
        printPreviewControl.Document = printDocument;
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
        //toolStripStatusLabel.Text = $"Format: {printDocument.DefaultPageSettings.PaperSize.PaperName}";
        if (printDocument.PrinterSettings.IsValid && printDocument.DefaultPageSettings.PaperSize != null)
        {
            toolStripStatusLabel.Text = $"Format: {printDocument.DefaultPageSettings.PaperSize.PaperName}";
        }
        else { toolStripStatusLabel.Text = "Kein gültiger Drucker ausgewählt."; }
    }

    private void PicPortrait_Click(object sender, EventArgs e) => rbPortrait.Checked = true;
    private void PicLandscape_Click(object sender, EventArgs e) => rbLandscape.Checked = true;


    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        switch (keyData)
        {
            case Keys.Tab:
                tabControl.SelectedIndex = (tabControl.SelectedIndex + 1) % tabControl.TabCount;
                return true;
            case Keys.Oemplus | Keys.Control:
            case Keys.Add | Keys.Control:
                if (printPreviewControl.Zoom < 1D) { printPreviewControl.Zoom += 0.1; if (printPreviewControl.Zoom > 1) { printPreviewControl.Zoom = 1D; } }
                return true;
            case Keys.OemMinus | Keys.Control:
            case Keys.Subtract | Keys.Control:
                if (printPreviewControl.Zoom >= 0.3D) { printPreviewControl.Zoom -= 0.1; }
                return true;
            case Keys.NumPad0 | Keys.Control:
            case Keys.D0 | Keys.Control:
                printPreviewControl.Zoom = _zoom;
                return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        if (printPreviewControl.Focus() && (ModifierKeys & Keys.Control) != 0)
        {
            var zoom = printPreviewControl.Zoom *= e.Delta > 0 ? 1.1 : 0.9;
            printPreviewControl.Zoom = Math.Min(1, Math.Max(.2, zoom));
        }
        base.OnMouseWheel(e);
    }

    private void CkbAnredePrint_CheckedChanged(object sender, EventArgs e)
    {
        if (ckbAnredePrint.Visible && ckbAnredePrint.Focused) { printPreviewControl.Document = printDocument; }
    }

    private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tabControl.SelectedIndex == 0)
        {
            toolStripStatusLabel.Text = $"Format: {printDocument.DefaultPageSettings.PaperSize.PaperName}";
        }
        else if (tabControl.SelectedIndex == 1)
        {
            toolStripStatusLabel.Text = $"Drucker: {printDocument.PrinterSettings.PrinterName}";
        }
        else
        {
            toolStripStatusLabel.Text = $"{printDocument.PrinterSettings.PrinterName}, {printDocument.DefaultPageSettings.PaperSize.PaperName}";
        }
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
        //toolStripStatusLabel.Text = $"Format: {printDocument.DefaultPageSettings.PaperSize.PaperName}";
    }

    //private void PrintPreviewControl_Paint(object sender, PaintEventArgs e) =>  e.Graphics.DrawString($"Zoom {_zoom * 100}%", new Font("Arial", 16, FontStyle.Bold), Brushes.Red, new PointF(10, 10));
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

    private void CkbPrintSender_CheckedChanged(object sender, EventArgs e)
    {
        if (ckbPrintSender.Visible && ckbPrintSender.Focused) { printPreviewControl.Document = printDocument; }
    }

    private void TcSender_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tcSender.Visible && tcSender.Focused) { printPreviewControl.Document = printDocument; }
    }

    private void NudOffset_ValueChanged(object sender, EventArgs e)
    {
        if (sender is NumericUpDown nud && nud.Visible && nud.Focused) { printPreviewControl.Document = printDocument; }
    }

    //private void NudRecipOffsetX_ValueChanged(object sender, EventArgs e)
    //{
    //    if (nudRecipOffsetX.Visible && nudRecipOffsetX.Focused) { printPreviewControl.Document = printDocument; }
    //}

    //private void NudRecipOffsetY_ValueChanged(object sender, EventArgs e)
    //{
    //    if (nudRecipOffsetY.Visible && nudRecipOffsetY.Focused) { printPreviewControl.Document = printDocument; }
    //}

    //private void NudSenderOffsetX_ValueChanged(object sender, EventArgs e)
    //{
    //    if (nudSenderOffsetX.Visible && nudSenderOffsetX.Focused) { printPreviewControl.Document = printDocument; }
    //}

    //private void NudSenderOffsetY_ValueChanged(object sender, EventArgs e)
    //{
    //    if (nudSenderOffsetY.Visible && nudSenderOffsetY.Focused) { printPreviewControl.Document = printDocument; }
    //}

    private void CkbBoldRecipient_CheckedChanged(object sender, EventArgs e)
    {

    }

    private void CkbBoldSender_CheckedChanged(object sender, EventArgs e)
    {

    }
}



