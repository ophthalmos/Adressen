namespace Adressen;

public partial class FrmColumns : Form
{
    private readonly bool[] _defaultHideArr;

    // Konstruktor nimmt jetzt die aktuellen UND die Standard-Werte entgegen
    public FrmColumns(bool[] currentHideArr, bool[] defaultHideArr)
    {
        InitializeComponent();
        _defaultHideArr = defaultHideArr;

        var limit = Math.Min(listView.Items.Count, currentHideArr.Length);
        for (var i = 0; i < limit; i++)
        {
            listView.Items[i].Checked = !currentHideArr[i];
        }
    }

    private void BtnStandard_Click(object sender, EventArgs e)
    {
        // Standardwerte anwenden
        var limit = Math.Min(listView.Items.Count, _defaultHideArr.Length);
        for (var i = 0; i < limit; i++)
        {
            listView.Items[i].Checked = !_defaultHideArr[i];
        }
    }

    // Die Hauptform ruft nur noch diese Methode auf, um das saubere Endergebnis zu bekommen
    public bool[] GetNewVisibilityArray()
    {
        var itemCount = listView.Items.Count;
        var newArr = new bool[itemCount];
        for (var i = 0; i < itemCount; i++)
        {
            newArr[i] = !listView.Items[i].Checked;
        }
        return newArr;
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape)
        {
            Close();
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }
}