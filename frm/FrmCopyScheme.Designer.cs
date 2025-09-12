namespace Adressen;

partial class FrmCopyScheme
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
        tbPattern1 = new TextBox();
        cbxFields = new ComboBox();
        btnInsert = new Button();
        tabControl = new TabControl();
        tabPage1 = new TabPage();
        tabPage2 = new TabPage();
        tbPattern2 = new TextBox();
        tabPage3 = new TabPage();
        tbPattern3 = new TextBox();
        tabPage4 = new TabPage();
        tbPattern4 = new TextBox();
        tabPage5 = new TabPage();
        tbPattern5 = new TextBox();
        tabPage6 = new TabPage();
        tbPattern6 = new TextBox();
        btnCopy = new Button();
        tbResult = new TextBox();
        panel = new Panel();
        tabControl.SuspendLayout();
        tabPage1.SuspendLayout();
        tabPage2.SuspendLayout();
        tabPage3.SuspendLayout();
        tabPage4.SuspendLayout();
        tabPage5.SuspendLayout();
        tabPage6.SuspendLayout();
        panel.SuspendLayout();
        SuspendLayout();
        // 
        // tbPattern1
        // 
        tbPattern1.AcceptsReturn = true;
        tbPattern1.AcceptsTab = true;
        tbPattern1.BackColor = Color.Ivory;
        tbPattern1.Dock = DockStyle.Fill;
        tbPattern1.Location = new Point(3, 3);
        tbPattern1.Multiline = true;
        tbPattern1.Name = "tbPattern1";
        tbPattern1.Size = new Size(250, 141);
        tbPattern1.TabIndex = 0;
        tbPattern1.WordWrap = false;
        tbPattern1.TextChanged += TbPattern_TextChanged;
        // 
        // cbxFields
        // 
        cbxFields.DropDownStyle = ComboBoxStyle.DropDownList;
        cbxFields.FormattingEnabled = true;
        cbxFields.Location = new Point(35, 166);
        cbxFields.Name = "cbxFields";
        cbxFields.Size = new Size(154, 25);
        cbxFields.TabIndex = 1;
        // 
        // btnInsert
        // 
        btnInsert.Location = new Point(195, 164);
        btnInsert.Name = "btnInsert";
        btnInsert.Size = new Size(90, 27);
        btnInsert.TabIndex = 2;
        btnInsert.Text = "Einfügen ⇑";
        btnInsert.UseVisualStyleBackColor = true;
        btnInsert.Click += BtnInsert_Click;
        // 
        // tabControl
        // 
        tabControl.Alignment = TabAlignment.Left;
        tabControl.Controls.Add(tabPage1);
        tabControl.Controls.Add(tabPage2);
        tabControl.Controls.Add(tabPage3);
        tabControl.Controls.Add(tabPage4);
        tabControl.Controls.Add(tabPage5);
        tabControl.Controls.Add(tabPage6);
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.ItemSize = new Size(25, 25);
        tabControl.Location = new Point(3, 3);
        tabControl.Multiline = true;
        tabControl.Name = "tabControl";
        tabControl.SelectedIndex = 0;
        tabControl.ShowToolTips = true;
        tabControl.Size = new Size(289, 155);
        tabControl.SizeMode = TabSizeMode.Fixed;
        tabControl.TabIndex = 3;
        tabControl.DrawItem += TabControl_DrawItem;
        tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;
        // 
        // tabPage1
        // 
        tabPage1.Controls.Add(tbPattern1);
        tabPage1.Location = new Point(29, 4);
        tabPage1.Name = "tabPage1";
        tabPage1.Padding = new Padding(3);
        tabPage1.Size = new Size(256, 147);
        tabPage1.TabIndex = 0;
        tabPage1.Text = "1";
        tabPage1.UseVisualStyleBackColor = true;
        // 
        // tabPage2
        // 
        tabPage2.Controls.Add(tbPattern2);
        tabPage2.Location = new Point(29, 4);
        tabPage2.Name = "tabPage2";
        tabPage2.Padding = new Padding(3);
        tabPage2.Size = new Size(256, 147);
        tabPage2.TabIndex = 1;
        tabPage2.Text = "2";
        tabPage2.UseVisualStyleBackColor = true;
        // 
        // tbPattern2
        // 
        tbPattern2.AcceptsReturn = true;
        tbPattern2.AcceptsTab = true;
        tbPattern2.BackColor = Color.Ivory;
        tbPattern2.Dock = DockStyle.Fill;
        tbPattern2.Location = new Point(3, 3);
        tbPattern2.Multiline = true;
        tbPattern2.Name = "tbPattern2";
        tbPattern2.Size = new Size(250, 141);
        tbPattern2.TabIndex = 1;
        tbPattern2.WordWrap = false;
        tbPattern2.TextChanged += TbPattern_TextChanged;
        // 
        // tabPage3
        // 
        tabPage3.Controls.Add(tbPattern3);
        tabPage3.Location = new Point(29, 4);
        tabPage3.Name = "tabPage3";
        tabPage3.Padding = new Padding(3);
        tabPage3.Size = new Size(256, 147);
        tabPage3.TabIndex = 2;
        tabPage3.Text = "3";
        tabPage3.UseVisualStyleBackColor = true;
        // 
        // tbPattern3
        // 
        tbPattern3.AcceptsReturn = true;
        tbPattern3.AcceptsTab = true;
        tbPattern3.BackColor = Color.Ivory;
        tbPattern3.Dock = DockStyle.Fill;
        tbPattern3.Location = new Point(3, 3);
        tbPattern3.Multiline = true;
        tbPattern3.Name = "tbPattern3";
        tbPattern3.Size = new Size(250, 141);
        tbPattern3.TabIndex = 1;
        tbPattern3.WordWrap = false;
        tbPattern3.TextChanged += TbPattern_TextChanged;
        // 
        // tabPage4
        // 
        tabPage4.Controls.Add(tbPattern4);
        tabPage4.Location = new Point(29, 4);
        tabPage4.Name = "tabPage4";
        tabPage4.Padding = new Padding(3);
        tabPage4.Size = new Size(256, 147);
        tabPage4.TabIndex = 3;
        tabPage4.Text = "4";
        tabPage4.UseVisualStyleBackColor = true;
        // 
        // tbPattern4
        // 
        tbPattern4.AcceptsReturn = true;
        tbPattern4.AcceptsTab = true;
        tbPattern4.BackColor = Color.Ivory;
        tbPattern4.Dock = DockStyle.Fill;
        tbPattern4.Location = new Point(3, 3);
        tbPattern4.Multiline = true;
        tbPattern4.Name = "tbPattern4";
        tbPattern4.Size = new Size(250, 141);
        tbPattern4.TabIndex = 1;
        tbPattern4.WordWrap = false;
        tbPattern4.TextChanged += TbPattern_TextChanged;
        // 
        // tabPage5
        // 
        tabPage5.Controls.Add(tbPattern5);
        tabPage5.Location = new Point(29, 4);
        tabPage5.Name = "tabPage5";
        tabPage5.Padding = new Padding(3);
        tabPage5.Size = new Size(256, 147);
        tabPage5.TabIndex = 4;
        tabPage5.Text = "5";
        tabPage5.UseVisualStyleBackColor = true;
        // 
        // tbPattern5
        // 
        tbPattern5.AcceptsReturn = true;
        tbPattern5.AcceptsTab = true;
        tbPattern5.BackColor = Color.Ivory;
        tbPattern5.Dock = DockStyle.Fill;
        tbPattern5.Location = new Point(3, 3);
        tbPattern5.Multiline = true;
        tbPattern5.Name = "tbPattern5";
        tbPattern5.Size = new Size(250, 141);
        tbPattern5.TabIndex = 1;
        tbPattern5.WordWrap = false;
        tbPattern5.TextChanged += TbPattern_TextChanged;
        // 
        // tabPage6
        // 
        tabPage6.Controls.Add(tbPattern6);
        tabPage6.Location = new Point(29, 4);
        tabPage6.Name = "tabPage6";
        tabPage6.Padding = new Padding(3);
        tabPage6.Size = new Size(256, 147);
        tabPage6.TabIndex = 5;
        tabPage6.Text = "6";
        tabPage6.UseVisualStyleBackColor = true;
        // 
        // tbPattern6
        // 
        tbPattern6.AcceptsReturn = true;
        tbPattern6.AcceptsTab = true;
        tbPattern6.BackColor = Color.Ivory;
        tbPattern6.Dock = DockStyle.Fill;
        tbPattern6.Location = new Point(3, 3);
        tbPattern6.Multiline = true;
        tbPattern6.Name = "tbPattern6";
        tbPattern6.Size = new Size(250, 141);
        tbPattern6.TabIndex = 1;
        tbPattern6.WordWrap = false;
        tbPattern6.TextChanged += TbPattern_TextChanged;
        // 
        // btnCopy
        // 
        btnCopy.DialogResult = DialogResult.OK;
        btnCopy.Image = Properties.Resources.clipboard_plus16;
        btnCopy.Location = new Point(300, 164);
        btnCopy.Name = "btnCopy";
        btnCopy.Size = new Size(250, 27);
        btnCopy.TabIndex = 5;
        btnCopy.Text = "Text in Zwischenablage kopieren";
        btnCopy.TextAlign = ContentAlignment.MiddleRight;
        btnCopy.TextImageRelation = TextImageRelation.ImageBeforeText;
        btnCopy.UseVisualStyleBackColor = true;
        btnCopy.Click += BtnCopy_Click;
        // 
        // tbResult
        // 
        tbResult.BackColor = Color.AliceBlue;
        tbResult.Location = new Point(300, 10);
        tbResult.Multiline = true;
        tbResult.Name = "tbResult";
        tbResult.ReadOnly = true;
        tbResult.Size = new Size(250, 141);
        tbResult.TabIndex = 0;
        tbResult.WordWrap = false;
        // 
        // panel
        // 
        panel.BackColor = SystemColors.ControlLightLight;
        panel.Controls.Add(btnInsert);
        panel.Controls.Add(cbxFields);
        panel.Controls.Add(tabControl);
        panel.Dock = DockStyle.Left;
        panel.Location = new Point(0, 0);
        panel.Name = "panel";
        panel.Size = new Size(297, 203);
        panel.TabIndex = 6;
        // 
        // FrmCopyScheme
        // 
        AcceptButton = btnCopy;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(562, 203);
        Controls.Add(btnCopy);
        Controls.Add(panel);
        Controls.Add(tbResult);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "FrmCopyScheme";
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Kopierschemata";
        Load += FrmCopyScheme_Load;
        Shown += FrmCopyScheme_Shown;
        tabControl.ResumeLayout(false);
        tabPage1.ResumeLayout(false);
        tabPage1.PerformLayout();
        tabPage2.ResumeLayout(false);
        tabPage2.PerformLayout();
        tabPage3.ResumeLayout(false);
        tabPage3.PerformLayout();
        tabPage4.ResumeLayout(false);
        tabPage4.PerformLayout();
        tabPage5.ResumeLayout(false);
        tabPage5.PerformLayout();
        tabPage6.ResumeLayout(false);
        tabPage6.PerformLayout();
        panel.ResumeLayout(false);
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private TextBox tbPattern1;
    private ComboBox cbxFields;
    private Button btnCopy;
    private Button btnInsert;
    private TextBox tbResult;
    private TabControl tabControl;
    private TabPage tabPage1;
    private TabPage tabPage2;
    private TabPage tabPage3;
    private TextBox tbPattern2;
    private TextBox tbPattern3;
    private Panel panel;
    private TabPage tabPage4;
    private TextBox tbPattern4;
    private TabPage tabPage5;
    private TextBox tbPattern5;
    private TabPage tabPage6;
    private TextBox tbPattern6;
}