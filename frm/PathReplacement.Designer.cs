namespace Adressen.frm;

partial class PathReplacement
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
        components = new System.ComponentModel.Container();
        var resources = new System.ComponentModel.ComponentResourceManager(typeof(PathReplacement));
        labelSearch = new Label();
        labelReplace = new Label();
        tbReplace = new TextBox();
        tbSearch = new TextBox();
        btnBrowse = new Button();
        btnOk = new Button();
        btnCancel = new Button();
        infoPictureBox = new PictureBox();
        helpToolTip = new ToolTip(components);
        ((System.ComponentModel.ISupportInitialize)infoPictureBox).BeginInit();
        SuspendLayout();
        // 
        // labelSearch
        // 
        labelSearch.AutoSize = true;
        labelSearch.Location = new Point(12, 15);
        labelSearch.Name = "labelSearch";
        labelSearch.Size = new Size(89, 19);
        labelSearch.TabIndex = 0;
        labelSearch.Text = "Suchen nach:";
        // 
        // labelReplace
        // 
        labelReplace.AutoSize = true;
        labelReplace.Location = new Point(12, 46);
        labelReplace.Name = "labelReplace";
        labelReplace.Size = new Size(102, 19);
        labelReplace.TabIndex = 1;
        labelReplace.Text = "Ersetzen durch:";
        // 
        // tbReplace
        // 
        tbReplace.Location = new Point(120, 44);
        tbReplace.Name = "tbReplace";
        tbReplace.Size = new Size(250, 25);
        tbReplace.TabIndex = 2;
        // 
        // tbSearch
        // 
        tbSearch.Location = new Point(120, 13);
        tbSearch.Name = "tbSearch";
        tbSearch.Size = new Size(250, 25);
        tbSearch.TabIndex = 3;
        // 
        // btnBrowse
        // 
        btnBrowse.Font = new Font("Segoe UI", 10F);
        btnBrowse.Location = new Point(376, 42);
        btnBrowse.Name = "btnBrowse";
        btnBrowse.Size = new Size(36, 25);
        btnBrowse.TabIndex = 6;
        btnBrowse.Text = "⚙";
        btnBrowse.UseVisualStyleBackColor = true;
        btnBrowse.Click += BtnBrowse_Click;
        // 
        // btnOk
        // 
        btnOk.DialogResult = DialogResult.OK;
        btnOk.Location = new Point(216, 82);
        btnOk.Name = "btnOk";
        btnOk.Size = new Size(95, 28);
        btnOk.TabIndex = 7;
        btnOk.Text = "&Ersetzen";
        btnOk.UseVisualStyleBackColor = true;
        // 
        // btnCancel
        // 
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new Point(317, 82);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(95, 28);
        btnCancel.TabIndex = 8;
        btnCancel.Text = "&Abbrechen";
        btnCancel.UseVisualStyleBackColor = true;
        // 
        // infoPictureBox
        // 
        infoPictureBox.Image = Properties.Resources.Help24;
        infoPictureBox.Location = new Point(381, 12);
        infoPictureBox.Name = "infoPictureBox";
        infoPictureBox.Size = new Size(24, 24);
        infoPictureBox.SizeMode = PictureBoxSizeMode.AutoSize;
        infoPictureBox.TabIndex = 9;
        infoPictureBox.TabStop = false;
        helpToolTip.SetToolTip(infoPictureBox, resources.GetString("infoPictureBox.ToolTip"));
        // 
        // helpToolTip
        // 
        helpToolTip.AutoPopDelay = 15000;
        helpToolTip.InitialDelay = 500;
        helpToolTip.IsBalloon = true;
        helpToolTip.ReshowDelay = 100;
        helpToolTip.ToolTipIcon = ToolTipIcon.Info;
        helpToolTip.ToolTipTitle = "Hilfe zur Pfad-Korrektur";
        // 
        // PathReplacement
        // 
        AcceptButton = btnOk;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = btnCancel;
        ClientSize = new Size(424, 122);
        Controls.Add(infoPictureBox);
        Controls.Add(btnCancel);
        Controls.Add(btnOk);
        Controls.Add(btnBrowse);
        Controls.Add(tbSearch);
        Controls.Add(tbReplace);
        Controls.Add(labelReplace);
        Controls.Add(labelSearch);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "PathReplacement";
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Pfadteile ersetzen";
        FormClosing += PathReplacement_FormClosing;
        ((System.ComponentModel.ISupportInitialize)infoPictureBox).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private Label labelSearch;
    private Label labelReplace;
    private TextBox tbReplace;
    private TextBox tbSearch;
    private Button btnBrowse;
    private Button btnOk;
    private Button btnCancel;
    private PictureBox infoPictureBox;
    private ToolTip helpToolTip;
}