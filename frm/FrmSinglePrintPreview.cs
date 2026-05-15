using System.Drawing.Printing;
using Adressen.Properties;
using System.ComponentModel;

namespace Adressen.frm;

public partial class FrmSinglePrintPreview : Form
{
    public record AddressPrintData
    {
        public string Title { get; init; } = "";
        public List<string> Groups { get; init; } = [];
        public List<(string Label, string Val)> NameFields { get; init; } = [];
        public List<(string Label, string Val)> AnschriftFields { get; init; } = [];
        public List<(string Label, string Val)> KommFields { get; init; } = [];
        public string Notes { get; init; } = "";
    }

    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public AddressPrintData? AddressData
    {
        get; set;
    }
    private bool _showNotes = true;
    private bool _largeFont = false;
    private string _paperSize = "DIN A4";

    public FrmSinglePrintPreview(int parentHeight)
    {
        InitializeComponent();
        Height = parentHeight;
        var menu = (ToolStripDropDownMenu)splitBtnPaperSize.DropDown;
        if (menu != null)
        {
            menu.ShowImageMargin = false;
            menu.ShowCheckMargin = false;
        }
    }

    private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
    {
        if (e.Graphics == null || AddressData == null) { return; }
        var g = e.Graphics;
        var margin = e.MarginBounds;
        var yPos = margin.Top;
        var baseSize = _largeFont ? 20 : 11;  // Schriftgrößen
        using var fontTitle = new Font("Segoe UI", 14, FontStyle.Bold);
        using var fontHeader = new Font("Segoe UI", 12, FontStyle.Bold);
        using var fontDate = new Font(fontHeader, FontStyle.Regular);
        using var fontRegular = new Font("Segoe UI", baseSize, FontStyle.Regular);
        using var brush = new SolidBrush(Color.Black);
        using var pen = new Pen(Color.LightGray, 1);

        var dateText = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
        g.DrawString(dateText, fontDate, brush, margin.Right - g.MeasureString(dateText, fontDate).Width, yPos);
        g.DrawString(AddressData.Title, fontTitle, brush, margin.Left, yPos);
        yPos += (int)fontTitle.GetHeight(g) + 30;

        void DrawSection(string sectionTitle, List<(string label, string val)> fields)  // Hilfsfuntion zum Zeichnen
        {
            var validFields = fields.Where(f => !string.IsNullOrWhiteSpace(f.val)).ToList();
            if (validFields.Count == 0) { return; }
            g.DrawString(sectionTitle, fontHeader, brush, margin.Left, yPos);
            yPos += (int)fontHeader.GetHeight(g) + 5;
            g.DrawLine(pen, margin.Left, yPos, margin.Right, yPos);
            yPos += 10;
            foreach (var (label, val) in validFields)
            {
                var valueX = margin.Left + (_largeFont ? 180 : 160);
                var rect = new RectangleF(valueX, yPos, margin.Right - valueX, margin.Bottom - yPos);
                g.DrawString(label + ":", fontRegular, brush, margin.Left, yPos);

                using var format = new StringFormat { Trimming = StringTrimming.Word };
                var size = g.MeasureString(val, fontRegular, (int)rect.Width, format);
                g.DrawString(val, fontRegular, brush, rect, format);
                yPos += (int)size.Height + 5;
            }
            yPos += 15;
        }

        DrawSection("Gruppen", [("Mitglied in", string.Join(", ", AddressData.Groups))]);
        DrawSection("Name", AddressData.NameFields);
        DrawSection("Anschrift", AddressData.AnschriftFields);
        DrawSection("Kommunikation", AddressData.KommFields);

        if (_showNotes && !string.IsNullOrEmpty(AddressData.Notes))
        {
            g.DrawString("Notizen", fontHeader, brush, margin.Left, yPos);
            yPos += (int)fontHeader.GetHeight(g) + 5;
            g.DrawLine(pen, margin.Left, yPos, margin.Right, yPos);
            yPos += 10;
            var rect = new RectangleF(margin.Left, yPos, margin.Width, margin.Bottom - yPos);
            g.DrawString(AddressData.Notes, fontRegular, brush, rect, new StringFormat { Trimming = StringTrimming.Word });
        }
        var footerText = "Adressen & Kontakte, www.netradio.info";
        using var footerFont = new Font("Segoe UI", 9, FontStyle.Regular);
        using var footerBrush = new SolidBrush(Color.DimGray);
        using var footerFormat = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Far // Richtet den Text am unteren Rand des Rechtecks aus
        };
        var footerRect = new RectangleF(margin.Left, margin.Top, margin.Width, margin.Height);
        g.DrawString(footerText, footerFont, footerBrush, footerRect, footerFormat);
        e.HasMorePages = false;
    }

    private void FrmSinglePrintPreview_Load(object sender, EventArgs e)
    {
        foreach (string printer in PrinterSettings.InstalledPrinters) { cbPrinter.Items.Add(printer); }  // Drucker laden
        cbPrinter.Text = printDocument.PrinterSettings.PrinterName;
        UpdatePaperSources();
        //printPreviewControl.Zoom = 1.0;
        printPreviewControl.AutoZoom = true;
    }

    private void PrintPreviewControl_Paint(object sender, PaintEventArgs e)
    {
        toolStripStatusLabel.Text = $"Zoom: {printPreviewControl.Zoom * 100:0}%{(printPreviewControl.AutoZoom ? " (Auto)" : "")}, Papiergröße: {_paperSize}";
    }

    private void UpdatePaperSources()
    {
        cbSources.Items.Clear();
        foreach (PaperSource source in printDocument.PrinterSettings.PaperSources) { cbSources.Items.Add(source.SourceName); }
        cbSources.SelectedIndex = 0;
    }

    private void CbPrinter_SelectedIndexChanged(object sender, EventArgs e)
    {
        printDocument.PrinterSettings.PrinterName = cbPrinter.Text;
        UpdatePaperSources();
        printPreviewControl.GeneratePreviewSilently();
    }

    private void CbSources_SelectedIndexChanged(object sender, EventArgs e)
    {
        var selectedSource = printDocument.PrinterSettings.PaperSources.Cast<PaperSource>().FirstOrDefault(s => s.SourceName == cbSources.Text);
        if (selectedSource != null) { printDocument.DefaultPageSettings.PaperSource = selectedSource; }
        printPreviewControl.GeneratePreviewSilently();
    }

    private void BtnFontSize_Click(object sender, EventArgs e)
    {
        _largeFont = !_largeFont;
        btnFontSize.Image = _largeFont ? Resources.DecreaseFontSize16 : Resources.IncreaseFontSize16;
        printPreviewControl.GeneratePreviewSilently();
    }

    private void BtnShowNotes_Click(object sender, EventArgs e)
    {
        _showNotes = !_showNotes;
        btnShowNotes.Image = _showNotes ? Resources.stickynoteminus16 : Resources.stickynoteplus16;
        printPreviewControl.GeneratePreviewSilently();
    }

    private void BtnPrint_Click(object? sender, EventArgs e)
    {
        printDocument.PrintController = new StandardPrintController();
        printDocument.Print();  // Ohne Dialog drucken
        Close();
    }

    private void BtnZoom_Click(object sender, EventArgs e)
    {
        if (printPreviewControl.AutoZoom)
        {
            printPreviewControl.AutoZoom = false;
            printPreviewControl.Zoom = 1.0;
        }
        else { printPreviewControl.AutoZoom = true; }
    }

    private void BtnClose_Click(object sender, EventArgs e) => Close();

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        switch (keyData)
        {
            case Keys.Escape:
                Close();
                return true;
            case Keys.Z | Keys.Control:
                BtnZoom_Click(null!, EventArgs.Empty);
                return true;
            case Keys.D | Keys.Control:
                var isA4 = printDocument.DefaultPageSettings.PaperSize.Kind == PaperKind.A4;
                var newKind = isA4 ? PaperKind.A5 : PaperKind.A4;
                SetPaperSizeAndMargins(newKind);
                return true;
            case Keys.G | Keys.Control:
                BtnFontSize_Click(null!, EventArgs.Empty);
                return true;
            case Keys.N | Keys.Control:
                BtnShowNotes_Click(null!, EventArgs.Empty);
                return true;
            case Keys.Enter:
            case Keys.P | Keys.Control:
                BtnPrint_Click(null, EventArgs.Empty);
                return true;
            case Keys.Oemplus | Keys.Control:
            case Keys.Add | Keys.Control:
                if (printPreviewControl.Zoom < 1.0)
                {
                    printPreviewControl.AutoZoom = false; // Explizit abschalten
                    printPreviewControl.Zoom = Math.Min(1.0, printPreviewControl.Zoom + 0.1);
                }
                else { Console.Beep(); }
                return true;
            case Keys.OemMinus | Keys.Control:
            case Keys.Subtract | Keys.Control:
                if (printPreviewControl.Zoom > 0.3)
                {
                    printPreviewControl.AutoZoom = false; // Explizit abschalten
                    printPreviewControl.Zoom = Math.Max(0.3, printPreviewControl.Zoom - 0.1);
                }
                else { Console.Beep(); }
                return true;
            case Keys.NumPad0 | Keys.Control:
            case Keys.D0 | Keys.Control:
                printPreviewControl.AutoZoom = true;
                return true;
            case Keys.NumPad4 | Keys.Control:
            case Keys.D4 | Keys.Control:
                SplitItemA4_Click(null!, EventArgs.Empty);
                return true;
            case Keys.NumPad5 | Keys.Control:
            case Keys.D5 | Keys.Control:
                SplitItemA5_Click(null!, EventArgs.Empty);
                return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void ToolStrip_Resize(object sender, EventArgs e) => btnShowNotes.DisplayStyle = btnFontSize.DisplayStyle = splitBtnPaperSize.DisplayStyle = 
        toolStrip.Width > 890 ? ToolStripItemDisplayStyle.ImageAndText : ToolStripItemDisplayStyle.Image;

    private void SetPaperSizeAndMargins(PaperKind kind)
    {
        var size = printDocument.PrinterSettings.PaperSizes.Cast<PaperSize>().FirstOrDefault(s => s.Kind == kind);
        if (size != null)
        {
            printDocument.DefaultPageSettings.PaperSize = size;
            _paperSize = kind == PaperKind.A4 ? "DIN A4" : kind == PaperKind.A5 ? "DIN A5" : "Unbekannt";
            if (kind == PaperKind.A5) { printDocument.DefaultPageSettings.Margins = new Margins(50, 50, 50, 50); }  // Schmalere Ränder für A5 (z.B. 0,5 Zoll ≈ 1,27 cm)
            else { printDocument.DefaultPageSettings.Margins = new Margins(100, 100, 100, 100); }  // Standardränder für A4 (1 Zoll ≈ 2,54 cm)
            printPreviewControl.GeneratePreviewSilently();
        }
        else { _paperSize = "Unbekannt"; }
    }

    private void SplitItemA4_Click(object sender, EventArgs e) => SetPaperSizeAndMargins(PaperKind.A4);
    private void SplitItemA5_Click(object sender, EventArgs e) => SetPaperSizeAndMargins(PaperKind.A5);

    private void SplitBtnPaperSize_ButtonClick(object sender, EventArgs e) => splitBtnPaperSize.ShowDropDown();
    private void CbPrinter_Click(object sender, EventArgs e) => cbPrinter.DroppedDown = true;
    private void CbSources_Click(object sender, EventArgs e) => cbSources.DroppedDown = true;

}
