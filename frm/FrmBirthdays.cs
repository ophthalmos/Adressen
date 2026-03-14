using System.Drawing.Drawing2D;
using Adressen.cls; // Namespace für AppSettings

namespace Adressen;

public partial class FrmBirthdays : Form
{
    // Diese Property bleibt, da sie reiner UI-Status ist (welche Zeile gewählt wurde)
    public int SelectionIndex => listView.SelectedIndices.Count > 0 ? listView.SelectedIndices[0] : -1;

    private readonly AppSettings _settings;
    private readonly List<int> _birthdayTodayList = [];
    private readonly Image partyHat = Properties.Resources.FavoriteStar16; // oder Ihr Bildname
    private readonly bool _isLocal; // Umbenannt in _isLocal (Naming Convention)
    private readonly int _initialIndex = -1;
    //private readonly StringFormat _sfNear = new()
    //{
    //    Alignment = StringAlignment.Near,
    //    LineAlignment = StringAlignment.Center,
    //    FormatFlags = StringFormatFlags.NoWrap,
    //    Trimming = StringTrimming.EllipsisCharacter
    //};

    // Konstruktor nimmt jetzt AppSettings direkt entgegen
    public FrmBirthdays(AppSettings settings, List<(DateOnly Datum, string Name, int Alter, int Tage, string Id)> geburtstage, bool localAdr)
    {
        InitializeComponent();
        typeof(Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.SetValue(listView, true, null);
        _settings = settings;
        _isLocal = localAdr;

        ApplyLVTextFont();
        // --- 1. Design & Farben ---
        ApplyColorScheme();

        // --- 2. Data Binding ---
        InitializeDataBindings();

        // --- 3. Listen-Logik (unverändert, aber sauberer) ---
        var nextBirthdayIndex = -1;
        var minDays = int.MaxValue;

        listView.BeginUpdate(); // Performance-Boost beim Füllen
        for (var i = 0; i < geburtstage.Count; i++)
        {
            var info = geburtstage[i];
            var item = new ListViewItem(info.Datum.ToShortDateString());
            item.SubItems.Add(info.Name);
            item.SubItems.Add(info.Alter.ToString());
            item.SubItems.Add(info.Tage.ToString());

            // Logik für "Nächsten Geburtstag" suchen
            if (info.Tage >= 0 && info.Tage < minDays)
            {
                minDays = info.Tage;
                nextBirthdayIndex = i;
            }

            if (info.Tage == 0) { _birthdayTodayList.Add(i); }
            listView.Items.Add(item);
        }
        listView.EndUpdate();

        // --- 4. Fokus setzen ---
        if (nextBirthdayIndex != -1)
        {
            _initialIndex = nextBirthdayIndex;
            AcceptButton = btnShowAddress;
        }
        else if (listView.Items.Count > 0)
        {
            AcceptButton = btnShowAddress;
            listView.Items[0].Selected = true;
            listView.EnsureVisible(0);
        }
    }

    private void ApplyColorScheme()
    {
        BackColor = _settings.ColorScheme switch
        {
            "blue" => SystemColors.GradientInactiveCaption,
            "pale" => SystemColors.ControlLightLight,
            "dark" => SystemColors.ControlDark,
            _ => SystemColors.Control,
        };
    }

    private void ApplyLVTextFont()
    {
        try
        {
            // Wir erstellen die Font basierend auf den Benutzereinstellungen
            var newFont = new Font(_settings.AppFontName, _settings.AppFontSize, FontStyle.Regular, GraphicsUnit.Point);

            // Die Zuweisung an die ListView sorgt dafür, dass DrawSubItem 
            // automatisch die richtige Font nutzt (via e.SubItem.Font oder listView.Font)
            listView.Font = newFont;
        }
        catch
        {
            // Fallback, falls die Schriftart im System Probleme macht
            listView.Font = new Font("Segoe UI", 10f);
        }
    }

    private void InitializeDataBindings()
    {
        // Numerische Felder direkt an Settings binden
        beforeNumUpDown.DataBindings.Add("Value", _settings, nameof(AppSettings.BirthdayRemindLimit), false, DataSourceUpdateMode.OnPropertyChanged);
        afterNumUpDown.DataBindings.Add("Value", _settings, nameof(AppSettings.BirthdayRemindAfter), false, DataSourceUpdateMode.OnPropertyChanged);

        // CheckBox dynamisch binden:
        // Wenn Adressen-Modus (isLocal) -> Binde an BirthdayAddressShow
        // Wenn Kontakte-Modus (!isLocal) -> Binde an BirthdayContactShow
        var targetProperty = _isLocal ? nameof(AppSettings.BirthdayAddressShow) : nameof(AppSettings.BirthdayContactShow);
        chkBxBirthdayAutoShow.DataBindings.Add("Checked", _settings, targetProperty, false, DataSourceUpdateMode.OnPropertyChanged);
    }

    // --- Event Handler ---

    private void ListView_SelectedIndexChanged(object sender, EventArgs e)
    {
        var hasSelection = listView.SelectedIndices.Count > 0;
        btnShowAddress.Enabled = hasSelection;
        AcceptButton = hasSelection ? btnShowAddress : btnCancel;
    }

    private void ListView_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
    {
        var g = e.Graphics;

        // 1. Hintergrund zeichnen
        var headerBackColor = Color.FromArgb(240, 240, 240);
        using (var backBrush = new SolidBrush(headerBackColor))
        {
            g.FillRectangle(backBrush, e.Bounds);
        }

        // 2. Rahmen zeichnen (Dezente Trennlinien)
        using (var pen = new Pen(SystemColors.ControlDark))
        {
            g.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
            g.DrawLine(pen, e.Bounds.Right - 1, e.Bounds.Top, e.Bounds.Right - 1, e.Bounds.Bottom - 1);
        }

        // 3. Text zeichnen mit TextRenderer (GDI statt GDI+) für scharfe Schrift
        if (e.Header != null)
        {
            var headerFont = Font;
            var textRect = e.Bounds;
            textRect.X += 4; // Padding
            textRect.Width -= 6;

            var flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine;

            // TextRenderer nutzt perfektes ClearType, wenn die Hintergrundfarbe mitgegeben wird
            TextRenderer.DrawText(g, e.Header.Text, headerFont, textRect, SystemColors.ControlText, headerBackColor, flags);
        }
    }

    private void ListView_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
    {
        if (e.Item == null || e.SubItem == null)
        {
            return;
        }

        var g = e.Graphics;

        // Antialiasing nur für das Bild aktivieren, für Text übernimmt das nun der TextRenderer
        g.SmoothingMode = SmoothingMode.HighQuality;

        var isSelected = e.Item.Selected;

        // 1. Hintergrund zeichnen
        var backColor = isSelected && _isLocal ? Color.FromArgb(176, 125, 71) : isSelected ? SystemColors.Highlight : e.SubItem.BackColor;
        using (var backBrush = new SolidBrush(backColor))
        {
            g.FillRectangle(backBrush, e.Bounds);
        }

        // 2. Text zeichnen mit TextRenderer
        var textColor = isSelected ? Color.White : e.SubItem.ForeColor;
        var textRect = e.Bounds;
        textRect.X += 4;
        textRect.Width -= 4;

        var flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine;

        // Durch die Übergabe von backColor wird die Kantenglättung (ClearType) optimal berechnet
        TextRenderer.DrawText(g, e.SubItem.Text, e.SubItem.Font, textRect, textColor, backColor, flags);

        // 3. Partyhut zeichnen (Spalte 1 = Name)
        if (e.ColumnIndex == 1 && _birthdayTodayList.Contains(e.ItemIndex))
        {
            var bildX = e.Bounds.Right - partyHat.Width - 4;
            var bildY = e.Bounds.Top + (e.Bounds.Height - partyHat.Height) / 2;
            g.DrawImage(partyHat, bildX, bildY);
        }
    }

    private void FrmBirthdays_Shown(object sender, EventArgs e)
    {
        AdjustNameColumnWidth();
        BringToFront(); // Sicherstellen, dass es vorne ist
        Activate();

        if (_initialIndex >= 0 && _initialIndex < listView.Items.Count)
        {
            listView.Focus();
            var item = listView.Items[_initialIndex];
            item.Selected = true;
            item.Focused = true;
            listView.EnsureVisible(_initialIndex);
        }
        else
        {
            listView.Focus();
        }
    }

    private void ListView_MouseDoubleClick(object sender, MouseEventArgs e) => DialogResult = DialogResult.OK;

    private void ListView_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Space && listView.SelectedIndices.Count > 0) { DialogResult = DialogResult.OK; }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape) { Close(); return true; }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void ListView_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
    {
        if (e.Item != null)
        {
            var rect = e.Item.Bounds;
            rect.Inflate(0, 1);  // Das Rechteck oben und unten um 1 Pixel erweitern, um hängen gebliebene Fokus-Linien sicher zu löschen.
            listView.Invalidate(rect);
        }
    }

    private void ListView_ClientSizeChanged(object sender, EventArgs e) => AdjustNameColumnWidth();

    private void AdjustNameColumnWidth()
    {
        // Sicherheitsabfrage: Wurden die Spalten schon geladen?
        if (listView.Columns.Count < 4)
        {
            return;
        }

        var fixedWidth = listView.Columns[0].Width + listView.Columns[2].Width + listView.Columns[3].Width;
        var availableWidth = listView.ClientSize.Width;
        var newWidth = availableWidth - fixedWidth;

        if (newWidth > 50)
        {
            listView.Columns[1].Width = newWidth;
        }
    }

}