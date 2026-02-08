namespace Adressen;

partial class FrmProgSettings
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
        var resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmProgSettings));
        tabControl = new TabControl();
        tpAllgemein = new TabPage();
        gbTextProcessing = new GroupBox();
        rbManualSelect = new RadioButton();
        rbLibreOffice = new RadioButton();
        rbMSWord = new RadioButton();
        gbxColorScheme = new GroupBox();
        rbtnPale = new RadioButton();
        rbtnDark = new RadioButton();
        rbtnBlue = new RadioButton();
        rbtnGrey = new RadioButton();
        tpAdressen = new TabPage();
        ckbAskBeforeDelete = new CheckBox();
        ckbAskBeforeSaveSQL = new CheckBox();
        groupBox = new GroupBox();
        btnStandardFile = new Button();
        tbStandard = new TextBox();
        rbStandard = new RadioButton();
        rbRecent = new RadioButton();
        rbEmpty = new RadioButton();
        gbDatabaseFolder = new GroupBox();
        btnDatabaseFolder = new Button();
        tbDatabaseFolder = new TextBox();
        tpKontakte = new TabPage();
        gbxContactsAutoload = new GroupBox();
        ckbContactsAutoload = new CheckBox();
        tpWatchFolder = new TabPage();
        lblWatcherInfo = new Label();
        lblWatchFolder = new Label();
        ckbWatchFolder = new CheckBox();
        btnWatchFolder = new Button();
        tbWatchFolder = new TextBox();
        tpSicherung = new TabPage();
        lblBackupFolder = new Label();
        btnExplorer = new Button();
        lblBackup = new Label();
        ckbBackup = new CheckBox();
        btnBackupFolder = new Button();
        tbBackupFolder = new TextBox();
        btnCancel = new Button();
        btnOK = new Button();
        folderBrowserDialog = new FolderBrowserDialog();
        openFileDialog = new OpenFileDialog();
        tabControl.SuspendLayout();
        tpAllgemein.SuspendLayout();
        gbTextProcessing.SuspendLayout();
        gbxColorScheme.SuspendLayout();
        tpAdressen.SuspendLayout();
        groupBox.SuspendLayout();
        gbDatabaseFolder.SuspendLayout();
        tpKontakte.SuspendLayout();
        gbxContactsAutoload.SuspendLayout();
        tpWatchFolder.SuspendLayout();
        tpSicherung.SuspendLayout();
        SuspendLayout();
        // 
        // tabControl
        // 
        tabControl.Alignment = TabAlignment.Left;
        tabControl.Controls.Add(tpAllgemein);
        tabControl.Controls.Add(tpAdressen);
        tabControl.Controls.Add(tpKontakte);
        tabControl.Controls.Add(tpWatchFolder);
        tabControl.Controls.Add(tpSicherung);
        tabControl.Dock = DockStyle.Top;
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.ItemSize = new Size(30, 110);
        tabControl.Location = new Point(0, 0);
        tabControl.Multiline = true;
        tabControl.Name = "tabControl";
        tabControl.SelectedIndex = 0;
        tabControl.Size = new Size(389, 312);
        tabControl.SizeMode = TabSizeMode.Fixed;
        tabControl.TabIndex = 0;
        tabControl.DrawItem += TabControl_DrawItem;
        // 
        // tpAllgemein
        // 
        tpAllgemein.BackColor = SystemColors.ControlLightLight;
        tpAllgemein.BorderStyle = BorderStyle.FixedSingle;
        tpAllgemein.Controls.Add(gbTextProcessing);
        tpAllgemein.Controls.Add(gbxColorScheme);
        tpAllgemein.Location = new Point(114, 4);
        tpAllgemein.Name = "tpAllgemein";
        tpAllgemein.Size = new Size(271, 304);
        tpAllgemein.TabIndex = 3;
        tpAllgemein.Text = " Allgemein";
        // 
        // gbTextProcessing
        // 
        gbTextProcessing.Controls.Add(rbManualSelect);
        gbTextProcessing.Controls.Add(rbLibreOffice);
        gbTextProcessing.Controls.Add(rbMSWord);
        gbTextProcessing.Location = new Point(3, 67);
        gbTextProcessing.Name = "gbTextProcessing";
        gbTextProcessing.Size = new Size(263, 113);
        gbTextProcessing.TabIndex = 6;
        gbTextProcessing.TabStop = false;
        gbTextProcessing.Text = "Textverarbeitungsprogramm";
        // 
        // rbManualSelect
        // 
        rbManualSelect.AutoSize = true;
        rbManualSelect.Location = new Point(30, 82);
        rbManualSelect.Name = "rbManualSelect";
        rbManualSelect.Size = new Size(150, 23);
        rbManualSelect.TabIndex = 2;
        rbManualSelect.Text = "Jedesmal auswählen";
        rbManualSelect.UseVisualStyleBackColor = true;
        // 
        // rbLibreOffice
        // 
        rbLibreOffice.AutoSize = true;
        rbLibreOffice.Location = new Point(30, 53);
        rbLibreOffice.Name = "rbLibreOffice";
        rbLibreOffice.Size = new Size(96, 23);
        rbLibreOffice.TabIndex = 1;
        rbLibreOffice.Text = "Libre Office";
        rbLibreOffice.UseVisualStyleBackColor = true;
        // 
        // rbMSWord
        // 
        rbMSWord.AutoSize = true;
        rbMSWord.Checked = true;
        rbMSWord.Location = new Point(30, 24);
        rbMSWord.Name = "rbMSWord";
        rbMSWord.Size = new Size(122, 23);
        rbMSWord.TabIndex = 0;
        rbMSWord.TabStop = true;
        rbMSWord.Text = "Microsoft Word";
        rbMSWord.UseVisualStyleBackColor = true;
        // 
        // gbxColorScheme
        // 
        gbxColorScheme.Controls.Add(rbtnPale);
        gbxColorScheme.Controls.Add(rbtnDark);
        gbxColorScheme.Controls.Add(rbtnBlue);
        gbxColorScheme.Controls.Add(rbtnGrey);
        gbxColorScheme.Location = new Point(3, 6);
        gbxColorScheme.Name = "gbxColorScheme";
        gbxColorScheme.Size = new Size(263, 55);
        gbxColorScheme.TabIndex = 5;
        gbxColorScheme.TabStop = false;
        gbxColorScheme.Text = "Farbschema";
        // 
        // rbtnPale
        // 
        rbtnPale.AutoSize = true;
        rbtnPale.Location = new Point(128, 24);
        rbtnPale.Name = "rbtnPale";
        rbtnPale.Size = new Size(57, 23);
        rbtnPale.TabIndex = 3;
        rbtnPale.TabStop = true;
        rbtnPale.Text = "Weiß";
        rbtnPale.UseVisualStyleBackColor = true;
        // 
        // rbtnDark
        // 
        rbtnDark.AutoSize = true;
        rbtnDark.Location = new Point(191, 24);
        rbtnDark.Name = "rbtnDark";
        rbtnDark.Size = new Size(70, 23);
        rbtnDark.TabIndex = 2;
        rbtnDark.TabStop = true;
        rbtnDark.Text = "Dunkel";
        rbtnDark.UseVisualStyleBackColor = true;
        // 
        // rbtnBlue
        // 
        rbtnBlue.AutoSize = true;
        rbtnBlue.Location = new Point(69, 24);
        rbtnBlue.Name = "rbtnBlue";
        rbtnBlue.Size = new Size(53, 23);
        rbtnBlue.TabIndex = 1;
        rbtnBlue.TabStop = true;
        rbtnBlue.Text = "Blau";
        rbtnBlue.UseVisualStyleBackColor = true;
        // 
        // rbtnGrey
        // 
        rbtnGrey.AutoSize = true;
        rbtnGrey.Checked = true;
        rbtnGrey.Location = new Point(6, 24);
        rbtnGrey.Name = "rbtnGrey";
        rbtnGrey.Size = new Size(57, 23);
        rbtnGrey.TabIndex = 0;
        rbtnGrey.TabStop = true;
        rbtnGrey.Text = "Grau";
        rbtnGrey.UseVisualStyleBackColor = true;
        // 
        // tpAdressen
        // 
        tpAdressen.BackColor = SystemColors.ControlLightLight;
        tpAdressen.BorderStyle = BorderStyle.FixedSingle;
        tpAdressen.Controls.Add(ckbAskBeforeDelete);
        tpAdressen.Controls.Add(ckbAskBeforeSaveSQL);
        tpAdressen.Controls.Add(groupBox);
        tpAdressen.Controls.Add(gbDatabaseFolder);
        tpAdressen.Location = new Point(114, 4);
        tpAdressen.Name = "tpAdressen";
        tpAdressen.Padding = new Padding(3);
        tpAdressen.Size = new Size(271, 304);
        tpAdressen.TabIndex = 0;
        tpAdressen.Text = " Lokale Adressen";
        // 
        // ckbAskBeforeDelete
        // 
        ckbAskBeforeDelete.AutoSize = true;
        ckbAskBeforeDelete.Checked = true;
        ckbAskBeforeDelete.CheckState = CheckState.Checked;
        ckbAskBeforeDelete.Location = new Point(6, 221);
        ckbAskBeforeDelete.Name = "ckbAskBeforeDelete";
        ckbAskBeforeDelete.Size = new Size(248, 23);
        ckbAskBeforeDelete.TabIndex = 5;
        ckbAskBeforeDelete.Text = "Sicherheitsabfrage vor dem Löschen";
        ckbAskBeforeDelete.UseVisualStyleBackColor = true;
        // 
        // ckbAskBeforeSaveSQL
        // 
        ckbAskBeforeSaveSQL.AutoSize = true;
        ckbAskBeforeSaveSQL.Location = new Point(6, 251);
        ckbAskBeforeSaveSQL.Name = "ckbAskBeforeSaveSQL";
        ckbAskBeforeSaveSQL.Size = new Size(249, 23);
        ckbAskBeforeSaveSQL.TabIndex = 4;
        ckbAskBeforeSaveSQL.Text = "Abfrage vor Datenbankspeicherung ";
        ckbAskBeforeSaveSQL.UseVisualStyleBackColor = true;
        // 
        // groupBox
        // 
        groupBox.Controls.Add(btnStandardFile);
        groupBox.Controls.Add(tbStandard);
        groupBox.Controls.Add(rbStandard);
        groupBox.Controls.Add(rbRecent);
        groupBox.Controls.Add(rbEmpty);
        groupBox.Location = new Point(6, 6);
        groupBox.Name = "groupBox";
        groupBox.Size = new Size(257, 137);
        groupBox.TabIndex = 2;
        groupBox.TabStop = false;
        groupBox.Text = "Lade bei Start des Programms";
        // 
        // btnStandardFile
        // 
        btnStandardFile.Enabled = false;
        btnStandardFile.Location = new Point(215, 101);
        btnStandardFile.Name = "btnStandardFile";
        btnStandardFile.Size = new Size(36, 25);
        btnStandardFile.TabIndex = 4;
        btnStandardFile.Text = "⚙";
        btnStandardFile.UseVisualStyleBackColor = true;
        btnStandardFile.Click += BtnStandardFile_Click;
        // 
        // tbStandard
        // 
        tbStandard.Enabled = false;
        tbStandard.Location = new Point(6, 102);
        tbStandard.Name = "tbStandard";
        tbStandard.Size = new Size(203, 25);
        tbStandard.TabIndex = 3;
        // 
        // rbStandard
        // 
        rbStandard.AutoSize = true;
        rbStandard.Location = new Point(10, 76);
        rbStandard.Name = "rbStandard";
        rbStandard.Size = new Size(206, 23);
        rbStandard.TabIndex = 2;
        rbStandard.TabStop = true;
        rbStandard.Text = "die folgende Datenbankdatei:";
        rbStandard.UseVisualStyleBackColor = true;
        rbStandard.CheckedChanged += RbStandard_CheckedChanged;
        // 
        // rbRecent
        // 
        rbRecent.AutoSize = true;
        rbRecent.Location = new Point(10, 50);
        rbRecent.Name = "rbRecent";
        rbRecent.Size = new Size(235, 23);
        rbRecent.TabIndex = 1;
        rbRecent.TabStop = true;
        rbRecent.Text = "die zuletzt verwendete Datenbank";
        rbRecent.UseVisualStyleBackColor = true;
        // 
        // rbEmpty
        // 
        rbEmpty.AutoSize = true;
        rbEmpty.Location = new Point(10, 24);
        rbEmpty.Name = "rbEmpty";
        rbEmpty.Size = new Size(149, 23);
        rbEmpty.TabIndex = 0;
        rbEmpty.TabStop = true;
        rbEmpty.Text = "keine Adressendatei";
        rbEmpty.UseVisualStyleBackColor = true;
        // 
        // gbDatabaseFolder
        // 
        gbDatabaseFolder.Controls.Add(btnDatabaseFolder);
        gbDatabaseFolder.Controls.Add(tbDatabaseFolder);
        gbDatabaseFolder.Location = new Point(6, 149);
        gbDatabaseFolder.Name = "gbDatabaseFolder";
        gbDatabaseFolder.Size = new Size(257, 56);
        gbDatabaseFolder.TabIndex = 0;
        gbDatabaseFolder.TabStop = false;
        gbDatabaseFolder.Text = "Standard-Datenbankordner";
        // 
        // btnDatabaseFolder
        // 
        btnDatabaseFolder.Location = new Point(215, 21);
        btnDatabaseFolder.Name = "btnDatabaseFolder";
        btnDatabaseFolder.Size = new Size(36, 25);
        btnDatabaseFolder.TabIndex = 1;
        btnDatabaseFolder.Text = "⚙";
        btnDatabaseFolder.UseVisualStyleBackColor = true;
        btnDatabaseFolder.Click += BtnDatabaseFolder_Click;
        // 
        // tbDatabaseFolder
        // 
        tbDatabaseFolder.Location = new Point(6, 21);
        tbDatabaseFolder.Name = "tbDatabaseFolder";
        tbDatabaseFolder.Size = new Size(203, 25);
        tbDatabaseFolder.TabIndex = 0;
        // 
        // tpKontakte
        // 
        tpKontakte.BackColor = SystemColors.ControlLightLight;
        tpKontakte.BorderStyle = BorderStyle.FixedSingle;
        tpKontakte.Controls.Add(gbxContactsAutoload);
        tpKontakte.Location = new Point(114, 4);
        tpKontakte.Name = "tpKontakte";
        tpKontakte.Padding = new Padding(3);
        tpKontakte.Size = new Size(271, 304);
        tpKontakte.TabIndex = 1;
        tpKontakte.Text = " Google Kontakte";
        // 
        // gbxContactsAutoload
        // 
        gbxContactsAutoload.Controls.Add(ckbContactsAutoload);
        gbxContactsAutoload.Location = new Point(6, 7);
        gbxContactsAutoload.Name = "gbxContactsAutoload";
        gbxContactsAutoload.Size = new Size(257, 55);
        gbxContactsAutoload.TabIndex = 6;
        gbxContactsAutoload.TabStop = false;
        gbxContactsAutoload.Text = "Autostart";
        // 
        // ckbContactsAutoload
        // 
        ckbContactsAutoload.AutoSize = true;
        ckbContactsAutoload.Location = new Point(6, 24);
        ckbContactsAutoload.Name = "ckbContactsAutoload";
        ckbContactsAutoload.Size = new Size(239, 23);
        ckbContactsAutoload.TabIndex = 2;
        ckbContactsAutoload.Text = "Kontakte bei Programmstart laden";
        ckbContactsAutoload.UseVisualStyleBackColor = true;
        // 
        // tpWatchFolder
        // 
        tpWatchFolder.BackColor = SystemColors.ControlLightLight;
        tpWatchFolder.BorderStyle = BorderStyle.FixedSingle;
        tpWatchFolder.Controls.Add(lblWatcherInfo);
        tpWatchFolder.Controls.Add(lblWatchFolder);
        tpWatchFolder.Controls.Add(ckbWatchFolder);
        tpWatchFolder.Controls.Add(btnWatchFolder);
        tpWatchFolder.Controls.Add(tbWatchFolder);
        tpWatchFolder.Location = new Point(114, 4);
        tpWatchFolder.Name = "tpWatchFolder";
        tpWatchFolder.Size = new Size(271, 304);
        tpWatchFolder.TabIndex = 4;
        tpWatchFolder.Text = " Briefzuordnung";
        // 
        // lblWatcherInfo
        // 
        lblWatcherInfo.Location = new Point(12, 14);
        lblWatcherInfo.Name = "lblWatcherInfo";
        lblWatcherInfo.Size = new Size(250, 100);
        lblWatcherInfo.TabIndex = 13;
        lblWatcherInfo.Text = "Beim Hinzufügen oder Ändern von Do-\r\nkumenten im Briefordner kann eine au-\r\ntomatische Benachrichtigung mit dem\r\nAngebot erfolgen, den Dateipfad in die\r\nBriefe-Liste der Adresse aufzunehmen.";
        // 
        // lblWatchFolder
        // 
        lblWatchFolder.AutoSize = true;
        lblWatchFolder.Location = new Point(12, 145);
        lblWatchFolder.Name = "lblWatchFolder";
        lblWatchFolder.Size = new Size(80, 19);
        lblWatchFolder.TabIndex = 12;
        lblWatchFolder.Text = "Briefordner:";
        // 
        // ckbWatchFolder
        // 
        ckbWatchFolder.AutoSize = true;
        ckbWatchFolder.Location = new Point(12, 119);
        ckbWatchFolder.Name = "ckbWatchFolder";
        ckbWatchFolder.Size = new Size(225, 23);
        ckbWatchFolder.TabIndex = 11;
        ckbWatchFolder.Text = "Auf Veränderungen überwachen";
        ckbWatchFolder.UseVisualStyleBackColor = true;
        ckbWatchFolder.CheckedChanged += CkbWatchFolder_CheckedChanged;
        // 
        // btnWatchFolder
        // 
        btnWatchFolder.Location = new Point(220, 167);
        btnWatchFolder.Name = "btnWatchFolder";
        btnWatchFolder.Size = new Size(36, 25);
        btnWatchFolder.TabIndex = 10;
        btnWatchFolder.Text = "⚙";
        btnWatchFolder.UseVisualStyleBackColor = true;
        btnWatchFolder.Click += BtnWatchFolder_Click;
        // 
        // tbWatchFolder
        // 
        tbWatchFolder.Location = new Point(12, 167);
        tbWatchFolder.Name = "tbWatchFolder";
        tbWatchFolder.Size = new Size(202, 25);
        tbWatchFolder.TabIndex = 9;
        // 
        // tpSicherung
        // 
        tpSicherung.BackColor = SystemColors.ControlLightLight;
        tpSicherung.BorderStyle = BorderStyle.FixedSingle;
        tpSicherung.Controls.Add(lblBackupFolder);
        tpSicherung.Controls.Add(btnExplorer);
        tpSicherung.Controls.Add(lblBackup);
        tpSicherung.Controls.Add(ckbBackup);
        tpSicherung.Controls.Add(btnBackupFolder);
        tpSicherung.Controls.Add(tbBackupFolder);
        tpSicherung.Location = new Point(114, 4);
        tpSicherung.Name = "tpSicherung";
        tpSicherung.Size = new Size(271, 304);
        tpSicherung.TabIndex = 2;
        tpSicherung.Text = " Sicherung";
        // 
        // lblBackupFolder
        // 
        lblBackupFolder.AutoSize = true;
        lblBackupFolder.Location = new Point(12, 34);
        lblBackupFolder.Name = "lblBackupFolder";
        lblBackupFolder.Size = new Size(119, 19);
        lblBackupFolder.TabIndex = 8;
        lblBackupFolder.Text = "Sicherungsordner:";
        // 
        // btnExplorer
        // 
        btnExplorer.Location = new Point(12, 265);
        btnExplorer.Name = "btnExplorer";
        btnExplorer.Size = new Size(244, 26);
        btnExplorer.TabIndex = 7;
        btnExplorer.Text = "Sicherungsordner anzeigen";
        btnExplorer.UseVisualStyleBackColor = true;
        btnExplorer.Click += BtnExplorer_Click;
        // 
        // lblBackup
        // 
        lblBackup.Location = new Point(12, 84);
        lblBackup.Name = "lblBackup";
        lblBackup.Size = new Size(250, 136);
        lblBackup.TabIndex = 6;
        lblBackup.Text = resources.GetString("lblBackup.Text");
        // 
        // ckbBackup
        // 
        ckbBackup.AutoSize = true;
        ckbBackup.Location = new Point(12, 8);
        ckbBackup.Name = "ckbBackup";
        ckbBackup.Size = new Size(235, 23);
        ckbBackup.TabIndex = 2;
        ckbBackup.Text = "Daten täglich automatisch sichern";
        ckbBackup.UseVisualStyleBackColor = true;
        ckbBackup.CheckedChanged += CkbBackup_CheckedChanged;
        // 
        // btnBackupFolder
        // 
        btnBackupFolder.Location = new Point(220, 56);
        btnBackupFolder.Name = "btnBackupFolder";
        btnBackupFolder.Size = new Size(36, 25);
        btnBackupFolder.TabIndex = 1;
        btnBackupFolder.Text = "⚙";
        btnBackupFolder.UseVisualStyleBackColor = true;
        btnBackupFolder.Click += BtnBackupFolder_Click;
        // 
        // tbBackupFolder
        // 
        tbBackupFolder.Location = new Point(12, 56);
        tbBackupFolder.Name = "tbBackupFolder";
        tbBackupFolder.Size = new Size(202, 25);
        tbBackupFolder.TabIndex = 0;
        tbBackupFolder.TextChanged += TbBackupFolder_TextChanged;
        // 
        // btnCancel
        // 
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new Point(287, 318);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(98, 26);
        btnCancel.TabIndex = 1;
        btnCancel.Text = "Abbrechen";
        btnCancel.UseVisualStyleBackColor = true;
        // 
        // btnOK
        // 
        btnOK.DialogResult = DialogResult.OK;
        btnOK.Location = new Point(114, 318);
        btnOK.Name = "btnOK";
        btnOK.Size = new Size(167, 26);
        btnOK.TabIndex = 2;
        btnOK.Text = "Einstellungen speichern";
        btnOK.UseVisualStyleBackColor = true;
        // 
        // folderBrowserDialog
        // 
        folderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer;
        // 
        // openFileDialog
        // 
        openFileDialog.DefaultExt = "adb";
        openFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";
        // 
        // FrmProgSettings
        // 
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(389, 353);
        Controls.Add(btnOK);
        Controls.Add(btnCancel);
        Controls.Add(tabControl);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "FrmProgSettings";
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Programmeinstellungen";
        FormClosing += FrmProgSettings_FormClosing;
        Load += FrmProgSettings_Load;
        tabControl.ResumeLayout(false);
        tpAllgemein.ResumeLayout(false);
        gbTextProcessing.ResumeLayout(false);
        gbTextProcessing.PerformLayout();
        gbxColorScheme.ResumeLayout(false);
        gbxColorScheme.PerformLayout();
        tpAdressen.ResumeLayout(false);
        tpAdressen.PerformLayout();
        groupBox.ResumeLayout(false);
        groupBox.PerformLayout();
        gbDatabaseFolder.ResumeLayout(false);
        gbDatabaseFolder.PerformLayout();
        tpKontakte.ResumeLayout(false);
        gbxContactsAutoload.ResumeLayout(false);
        gbxContactsAutoload.PerformLayout();
        tpWatchFolder.ResumeLayout(false);
        tpWatchFolder.PerformLayout();
        tpSicherung.ResumeLayout(false);
        tpSicherung.PerformLayout();
        ResumeLayout(false);
    }

    #endregion

    private TabControl tabControl;
    private TabPage tpAdressen;
    private TabPage tpKontakte;
    private Button btnCancel;
    private Button btnOK;
    private GroupBox gbDatabaseFolder;
    private Button btnDatabaseFolder;
    private TextBox tbDatabaseFolder;
    private TabPage tpSicherung;
    private FolderBrowserDialog folderBrowserDialog;
    private Button btnBackupFolder;
    private TextBox tbBackupFolder;
    private CheckBox ckbContactsAutoload;
    private TabPage tpAllgemein;
    private CheckBox ckbBackup;
    private GroupBox groupBox;
    private Button btnStandardFile;
    private TextBox tbStandard;
    private RadioButton rbStandard;
    private RadioButton rbRecent;
    private RadioButton rbEmpty;
    private OpenFileDialog openFileDialog;
    private Label lblBackup;
    private Button btnExplorer;
    private Label lblBackupFolder;
    private GroupBox gbxColorScheme;
    private RadioButton rbtnDark;
    private RadioButton rbtnBlue;
    private RadioButton rbtnGrey;
    private RadioButton rbtnPale;
    private GroupBox gbTextProcessing;
    private RadioButton rbManualSelect;
    private RadioButton rbLibreOffice;
    private RadioButton rbMSWord;
    private GroupBox gbxContactsAutoload;
    private CheckBox ckbAskBeforeSaveSQL;
    private CheckBox ckbAskBeforeDelete;
    private TabPage tpWatchFolder;
    private Label lblWatchFolder;
    private CheckBox ckbWatchFolder;
    private Button btnWatchFolder;
    private TextBox tbWatchFolder;
    private Label lblWatcherInfo;
}