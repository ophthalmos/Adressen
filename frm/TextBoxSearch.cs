using Adressen.cls;

namespace Adressen;

public partial class TextBoxSearch : Form
{
    public string SearchText => cbxSearch.Text;
    public bool MatchCase => checkCase.Checked;

    public TextBoxSearch(string searchString, bool caseChecked)
    {
        InitializeComponent();

        checkCase.Checked = caseChecked;
        cbxSearch.Items.Clear();
        cbxSearch.Text = searchString;
        if (TextBoxSearchManager.SearchHistory.Count > 0) { cbxSearch.Items.AddRange([.. TextBoxSearchManager.SearchHistory]); }
    }

    private void BtnSearch_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(cbxSearch.Text)) { return; }
        DialogResult = DialogResult.OK;
        Close();
    }
}