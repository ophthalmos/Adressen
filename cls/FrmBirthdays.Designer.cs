namespace Adressen;

partial class FrmBirthdays
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
        listView = new ListView();
        chDate = new ColumnHeader();
        chName = new ColumnHeader();
        chAge = new ColumnHeader();
        chSpan = new ColumnHeader();
        numericUpDown = new NumericUpDown();
        label = new Label();
        btnShowAddress = new Button();
        btnCancel = new Button();
        ((System.ComponentModel.ISupportInitialize)numericUpDown).BeginInit();
        SuspendLayout();
        // 
        // listView
        // 
        listView.Columns.AddRange(new ColumnHeader[] { chDate, chName, chAge, chSpan });
        listView.Dock = DockStyle.Top;
        listView.FullRowSelect = true;
        listView.HeaderStyle = ColumnHeaderStyle.Nonclickable;
        listView.Location = new Point(0, 0);
        listView.Name = "listView";
        listView.OwnerDraw = true;
        listView.Size = new Size(409, 218);
        listView.TabIndex = 0;
        listView.UseCompatibleStateImageBehavior = false;
        listView.View = View.Details;
        listView.DrawColumnHeader += ListView_DrawColumnHeader;
        listView.DrawSubItem += ListView_DrawSubItem;
        listView.SelectedIndexChanged += ListView_SelectedIndexChanged;
        listView.KeyDown += ListView_KeyDown;
        listView.MouseDoubleClick += ListView_MouseDoubleClick;
        // 
        // chDate
        // 
        chDate.Text = "Geburtstag";
        chDate.Width = 90;
        // 
        // chName
        // 
        chName.Text = "Name";
        chName.Width = 196;
        // 
        // chAge
        // 
        chAge.Text = "Alter";
        chAge.Width = 52;
        // 
        // chSpan
        // 
        chSpan.Text = "Noch";
        chSpan.Width = 50;
        // 
        // numericUpDown
        // 
        numericUpDown.Location = new Point(352, 224);
        numericUpDown.Maximum = new decimal(new int[] { 99, 0, 0, 0 });
        numericUpDown.Name = "numericUpDown";
        numericUpDown.Size = new Size(45, 25);
        numericUpDown.TabIndex = 1;
        numericUpDown.TextAlign = HorizontalAlignment.Center;
        // 
        // label
        // 
        label.AutoSize = true;
        label.Location = new Point(5, 226);
        label.Name = "label";
        label.Size = new Size(344, 19);
        label.TabIndex = 2;
        label.Text = "Maximale Anzahl der Tage vor dem Geburtstagstermin:";
        // 
        // btnShowAddress
        // 
        btnShowAddress.DialogResult = DialogResult.OK;
        btnShowAddress.Enabled = false;
        btnShowAddress.Location = new Point(117, 255);
        btnShowAddress.Name = "btnShowAddress";
        btnShowAddress.Size = new Size(164, 27);
        btnShowAddress.TabIndex = 3;
        btnShowAddress.Text = "Gehe zu Adresse";
        btnShowAddress.UseVisualStyleBackColor = true;
        // 
        // btnCancel
        // 
        btnCancel.DialogResult = DialogResult.Continue;
        btnCancel.Location = new Point(287, 255);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(110, 27);
        btnCancel.TabIndex = 4;
        btnCancel.Text = "Schließen";
        btnCancel.UseVisualStyleBackColor = true;
        // 
        // FrmBirthdays
        // 
        AcceptButton = btnCancel;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(409, 294);
        Controls.Add(btnCancel);
        Controls.Add(btnShowAddress);
        Controls.Add(label);
        Controls.Add(numericUpDown);
        Controls.Add(listView);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "FrmBirthdays";
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Anstehende Geburtstage";
        Shown += FrmBirthdays_Shown;
        ((System.ComponentModel.ISupportInitialize)numericUpDown).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private ListView listView;
    private NumericUpDown numericUpDown;
    private Label label;
    private Button btnShowAddress;
    private Button btnCancel;
    private ColumnHeader chDate;
    private ColumnHeader chName;
    private ColumnHeader chAge;
    private ColumnHeader chSpan;
}