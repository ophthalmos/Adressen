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
        beforeNumUpDown = new NumericUpDown();
        label1 = new Label();
        btnShowAddress = new Button();
        btnCancel = new Button();
        label2 = new Label();
        afterNumUpDown = new NumericUpDown();
        chkBxBirthdayAutoShow = new CheckBox();
        ((System.ComponentModel.ISupportInitialize)beforeNumUpDown).BeginInit();
        ((System.ComponentModel.ISupportInitialize)afterNumUpDown).BeginInit();
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
        listView.Size = new Size(409, 250);
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
        // beforeNumUpDown
        // 
        beforeNumUpDown.Location = new Point(352, 256);
        beforeNumUpDown.Maximum = new decimal(new int[] { 99, 0, 0, 0 });
        beforeNumUpDown.Name = "beforeNumUpDown";
        beforeNumUpDown.Size = new Size(45, 25);
        beforeNumUpDown.TabIndex = 1;
        beforeNumUpDown.TextAlign = HorizontalAlignment.Center;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(5, 258);
        label1.Name = "label1";
        label1.Size = new Size(344, 19);
        label1.TabIndex = 2;
        label1.Text = "Maximale Anzahl der Tage vor dem Geburtstagstermin:";
        // 
        // btnShowAddress
        // 
        btnShowAddress.DialogResult = DialogResult.OK;
        btnShowAddress.Enabled = false;
        btnShowAddress.Location = new Point(117, 318);
        btnShowAddress.Name = "btnShowAddress";
        btnShowAddress.Size = new Size(164, 27);
        btnShowAddress.TabIndex = 3;
        btnShowAddress.Text = "Gehe zu Adresse";
        btnShowAddress.UseVisualStyleBackColor = true;
        // 
        // btnCancel
        // 
        btnCancel.DialogResult = DialogResult.Continue;
        btnCancel.Location = new Point(287, 318);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(110, 27);
        btnCancel.TabIndex = 4;
        btnCancel.Text = "Schließen";
        btnCancel.UseVisualStyleBackColor = true;
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(5, 289);
        label2.Name = "label2";
        label2.Size = new Size(344, 19);
        label2.TabIndex = 6;
        label2.Text = "Tage nach den Geburtstag, die mit einbezogen werden:";
        // 
        // afterNumUpDown
        // 
        afterNumUpDown.Location = new Point(352, 287);
        afterNumUpDown.Maximum = new decimal(new int[] { 99, 0, 0, 0 });
        afterNumUpDown.Name = "afterNumUpDown";
        afterNumUpDown.Size = new Size(45, 25);
        afterNumUpDown.TabIndex = 5;
        afterNumUpDown.TextAlign = HorizontalAlignment.Center;
        // 
        // chkBxBirthdayAutoShow
        // 
        chkBxBirthdayAutoShow.AutoSize = true;
        chkBxBirthdayAutoShow.Location = new Point(10, 321);
        chkBxBirthdayAutoShow.Name = "chkBxBirthdayAutoShow";
        chkBxBirthdayAutoShow.Size = new Size(86, 23);
        chkBxBirthdayAutoShow.TabIndex = 7;
        chkBxBirthdayAutoShow.Text = "Autostart";
        chkBxBirthdayAutoShow.UseVisualStyleBackColor = true;
        // 
        // FrmBirthdays
        // 
        AcceptButton = btnCancel;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(409, 357);
        Controls.Add(chkBxBirthdayAutoShow);
        Controls.Add(label2);
        Controls.Add(afterNumUpDown);
        Controls.Add(btnCancel);
        Controls.Add(btnShowAddress);
        Controls.Add(label1);
        Controls.Add(beforeNumUpDown);
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
        ((System.ComponentModel.ISupportInitialize)beforeNumUpDown).EndInit();
        ((System.ComponentModel.ISupportInitialize)afterNumUpDown).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private ListView listView;
    private NumericUpDown beforeNumUpDown;
    private Label label1;
    private Button btnShowAddress;
    private Button btnCancel;
    private ColumnHeader chDate;
    private ColumnHeader chName;
    private ColumnHeader chAge;
    private ColumnHeader chSpan;
    private Label label2;
    private NumericUpDown afterNumUpDown;
    private CheckBox chkBxBirthdayAutoShow;
}