namespace Adressen.frm;
public partial class FrmGroupRename : Form
{
    public string GetText() => textBox.Text;

    public FrmGroupRename(string oldName)
    {
        InitializeComponent();
        textBox.Text = oldName; 
    }
}
