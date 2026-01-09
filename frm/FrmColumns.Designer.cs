namespace Adressen;

partial class FrmColumns
{
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        var listViewItem1 = new ListViewItem("Anrede");
        var listViewItem2 = new ListViewItem("Praefix");
        var listViewItem3 = new ListViewItem("Nachname");
        var listViewItem4 = new ListViewItem("Vorname");
        var listViewItem5 = new ListViewItem("Zwischenname");
        var listViewItem6 = new ListViewItem("Nickname");
        var listViewItem7 = new ListViewItem("Suffix");
        var listViewItem8 = new ListViewItem("Firma");
        var listViewItem9 = new ListViewItem("Strasse");
        var listViewItem10 = new ListViewItem("PLZ");
        var listViewItem11 = new ListViewItem("Ort");
        var listViewItem12 = new ListViewItem("Land");
        var listViewItem13 = new ListViewItem("Betreff");
        var listViewItem14 = new ListViewItem("Grussformel");
        var listViewItem15 = new ListViewItem("Schlussformel");
        var listViewItem16 = new ListViewItem("Geburtstag");
        var listViewItem17 = new ListViewItem("Mail1");
        var listViewItem18 = new ListViewItem("Mail2");
        var listViewItem19 = new ListViewItem("Telefon1");
        var listViewItem20 = new ListViewItem("Telefon2");
        var listViewItem21 = new ListViewItem("Mobil");
        var listViewItem22 = new ListViewItem("Fax");
        var listViewItem23 = new ListViewItem("Internet");
        var listViewItem24 = new ListViewItem("Notizen");
        var listViewItem25 = new ListViewItem("Id");
        listView = new ListView();
        columnHeader = new ColumnHeader();
        btnClose = new Button();
        btnStandard = new Button();
        SuspendLayout();
        // 
        // listView
        // 
        listView.CheckBoxes = true;
        listView.Columns.AddRange(new ColumnHeader[] { columnHeader });
        listView.FullRowSelect = true;
        listView.HeaderStyle = ColumnHeaderStyle.None;
        listViewItem1.StateImageIndex = 0;
        listViewItem2.StateImageIndex = 0;
        listViewItem3.StateImageIndex = 0;
        listViewItem4.StateImageIndex = 0;
        listViewItem5.StateImageIndex = 0;
        listViewItem6.StateImageIndex = 0;
        listViewItem7.StateImageIndex = 0;
        listViewItem8.StateImageIndex = 0;
        listViewItem9.StateImageIndex = 0;
        listViewItem10.StateImageIndex = 0;
        listViewItem11.StateImageIndex = 0;
        listViewItem12.StateImageIndex = 0;
        listViewItem13.StateImageIndex = 0;
        listViewItem14.StateImageIndex = 0;
        listViewItem15.StateImageIndex = 0;
        listViewItem16.StateImageIndex = 0;
        listViewItem17.StateImageIndex = 0;
        listViewItem18.StateImageIndex = 0;
        listViewItem19.StateImageIndex = 0;
        listViewItem20.StateImageIndex = 0;
        listViewItem21.StateImageIndex = 0;
        listViewItem22.StateImageIndex = 0;
        listViewItem23.StateImageIndex = 0;
        listViewItem24.StateImageIndex = 0;
        listViewItem25.StateImageIndex = 0;
        listView.Items.AddRange(new ListViewItem[] { listViewItem1, listViewItem2, listViewItem3, listViewItem4, listViewItem5, listViewItem6, listViewItem7, listViewItem8, listViewItem9, listViewItem10, listViewItem11, listViewItem12, listViewItem13, listViewItem14, listViewItem15, listViewItem16, listViewItem17, listViewItem18, listViewItem19, listViewItem20, listViewItem21, listViewItem22, listViewItem23, listViewItem24, listViewItem25 });
        listView.LabelWrap = false;
        listView.Location = new Point(12, 12);
        listView.MultiSelect = false;
        listView.Name = "listView";
        listView.ShowGroups = false;
        listView.Size = new Size(159, 529);
        listView.TabIndex = 0;
        listView.UseCompatibleStateImageBehavior = false;
        listView.View = View.Details;
        // 
        // columnHeader
        // 
        columnHeader.Width = 155;
        // 
        // btnClose
        // 
        btnClose.DialogResult = DialogResult.OK;
        btnClose.Location = new Point(93, 547);
        btnClose.Name = "btnClose";
        btnClose.Size = new Size(78, 26);
        btnClose.TabIndex = 1;
        btnClose.Text = "Schließen";
        btnClose.UseVisualStyleBackColor = true;
        // 
        // btnStandard
        // 
        btnStandard.Location = new Point(12, 547);
        btnStandard.Name = "btnStandard";
        btnStandard.Size = new Size(75, 26);
        btnStandard.TabIndex = 2;
        btnStandard.Text = "Standard";
        btnStandard.UseVisualStyleBackColor = true;
        btnStandard.Click += BtnStandard_Click;
        // 
        // FrmColumns
        // 
        AcceptButton = btnClose;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(183, 585);
        Controls.Add(btnStandard);
        Controls.Add(btnClose);
        Controls.Add(listView);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "FrmColumns";
        ShowIcon = false;
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Hide;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Spalten auwählen";
        ResumeLayout(false);
    }

    #endregion

    private ListView listView;
    private ColumnHeader columnHeader;
    private Button btnClose;
    private Button btnStandard;
}