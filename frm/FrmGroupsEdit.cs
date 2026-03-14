using System.Data;
using Adressen.cls;

namespace Adressen.frm;

public partial class FrmGroupsEdit : Form
{
    public Dictionary<string, string> groupNameMap = [];
    private record GroupItemData(string Name, int Count);

    public FrmGroupsEdit(Dictionary<string, int> groupDict)
    {
        InitializeComponent();
        listBox.ItemHeight = listBox.Font.Height + 4;
        var sortedGroups = groupDict.OrderByDescending(kvp => kvp.Key == "★").ThenByDescending(kvp => kvp.Value).ToList();
        groupNameMap = sortedGroups.ToDictionary(kvp => kvp.Key, kvp => kvp.Key);
        foreach (var kvp in sortedGroups) { listBox.Items.Add(new GroupItemData(kvp.Key, kvp.Value)); } // Daten direkt als Objekte hinzufügen
        UpdateStatusCount();
    }

    private void BtnDelete_Click(object sender, EventArgs e)
    {
        if (listBox.SelectedItem is GroupItemData selectedData)
        {
            if (selectedData.Name == "★") { return; }
            groupNameMap[selectedData.Name] = string.Empty;
            listBox.Items.Remove(selectedData); // Einfach das Objekt entfernen
            UpdateStatusCount();
            btnClose.Enabled = true;
            btnClose.Focus();
        }
    }

    private void BtnEdit_Click(object sender, EventArgs e)
    {
        if (listBox.SelectedItem is GroupItemData oldData)
        {
            if (oldData.Name == "★") { return; }
            using var frm = new FrmGroupRename(oldData.Name);
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                var newName = frm.GetText();
                if (!string.IsNullOrEmpty(newName))
                {
                    groupNameMap[oldData.Name] = newName; // 1. Dictionary aktualisieren
                    var index = listBox.SelectedIndex; // 2. Element in der ListBox austauschen
                    listBox.Items[index] = new GroupItemData(newName, oldData.Count);
                    UpdateStatusCount();
                    btnClose.Enabled = true;
                    btnClose.Focus();
                }
            }
        }
    }

    private void ListBox_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (listBox.SelectedItem is GroupItemData selectedData)
        {
            var isSpecialGroup = selectedData.Name == "★";
            btnEdit.Enabled = !isSpecialGroup;
            btnDelete.Enabled = !isSpecialGroup;
            return;
        }
        btnEdit.Enabled = btnDelete.Enabled = false;
    }

    private void FrmGroups_Shown(object sender, EventArgs e) => listBox.Focus();

    private void ListBox_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0)
        {
            return;
        }

        // Hintergrund zeichnen (behandelt Selektion automatisch)
        e.DrawBackground();

        if (listBox.Items[e.Index] is GroupItemData data)
        {
            var g = e.Graphics;

            // Das Geheimnis für scharfen Text
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            var fontToUse = e.Font ?? listBox.Font;
            var suffix = $" ({data.Count})";

            // Den Text berechnen (deine Utils-Funktion bleibt, sollte aber intern idealerweise auch auf GDI+ basieren)
            var displayText = Utils.TruncateMiddle(data.Name, suffix, fontToUse, e.Bounds.Width);

            // Text-Bereich mit Padding definieren
            var textBounds = new Rectangle(e.Bounds.X + 2, e.Bounds.Y, e.Bounds.Width - 4, e.Bounds.Height);

            using var textBrush = new SolidBrush(e.ForeColor);
            using var stringFormat = new StringFormat
            {
                Alignment = StringAlignment.Near,      // Links
                LineAlignment = StringAlignment.Center, // Vertikal mittig
                FormatFlags = StringFormatFlags.NoWrap,
                Trimming = StringTrimming.EllipsisCharacter
            };

            // Zeichnen mit GDI+
            g.DrawString(displayText, fontToUse, textBrush, textBounds, stringFormat);
        }

        // Fokus-Rechteck zeichnen (punktierte Linie bei Tastaturnavigation)
        e.DrawFocusRectangle();
    }

    private void UpdateStatusCount() => toolStripStatusLabel.Text = $"{listBox.Items.Count} Gruppen";

    private void ListBox_SizeChanged(object sender, EventArgs e) => listBox.Invalidate();
}
