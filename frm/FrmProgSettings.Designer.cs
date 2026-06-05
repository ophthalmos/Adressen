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
        if (_tabFont != null)
        {
            _tabFont.Dispose();
        }

        if (_tabStringFormat != null)
        {
            _tabStringFormat.Dispose();
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
        btnCancel = new Button();
        folderBrowserDialog = new FolderBrowserDialog();
        openFileDialog = new OpenFileDialog();
        btnOK = new Button();
        tpSicherung = new TabPage();
        gbBackupZip = new GroupBox();
        btnZipArchive = new Button();
        ckbZipArchive = new CheckBox();
        lblZipArchive = new Label();
        tbZipArchive = new TextBox();
        gbBackupDaily = new GroupBox();
        tbBackupFolder = new TextBox();
        lblBackupFolder = new Label();
        btnExplorer = new Button();
        lblBackup = new Label();
        ckbBackup = new CheckBox();
        btnBackupFolder = new Button();
        tpAskBefore = new TabPage();
        gbxMin2Tray = new GroupBox();
        ckbBalloonTipMin2Tray = new CheckBox();
        gbxAskEnvelope = new GroupBox();
        ckbAskPrintEnvelope = new CheckBox();
        gbxAskLocal = new GroupBox();
        ckbAskBeforeDelete = new CheckBox();
        ckbAskBeforeSaveSQLExpander = new CheckBox();
        ckbAskBeforeSaveSQL = new CheckBox();
        tpWatchFolder = new TabPage();
        gbWatchFolder = new GroupBox();
        ckbWatchFolder = new CheckBox();
        lblWatcherInfo = new Label();
        tbWatchFolder = new TextBox();
        btnWatchFolder = new Button();
        tpAutostart = new TabPage();
        gbBirthdayRemind = new GroupBox();
        lblBirtdayRemind = new Label();
        ckbBirthdayRemind = new CheckBox();
        groupBox1 = new GroupBox();
        ckbMin2Tray = new CheckBox();
        ckbAutostart = new CheckBox();
        lblAutostart = new Label();
        gbxContactsAutoload = new GroupBox();
        ckbContactsAutoload = new CheckBox();
        labelAutoAdressen = new Label();
        tpAdressen = new TabPage();
        lblToggleDatabase = new Label();
        groupBox = new GroupBox();
        btnStandardFile = new Button();
        tbStandard = new TextBox();
        rbStandard = new RadioButton();
        rbRecent = new RadioButton();
        rbEmpty = new RadioButton();
        gbDatabaseFolder = new GroupBox();
        btnDatabaseFolder = new Button();
        tbDatabaseFolder = new TextBox();
        tpAllgemein = new TabPage();
        gbxFontSize = new GroupBox();
        ckbPlaceholderText = new CheckBox();
        btnFontReset = new Button();
        nudFontSize = new NumericUpDown();
        cbxFontName = new ComboBox();
        gbTextProcessing = new GroupBox();
        rbManualSelect = new RadioButton();
        rbLibreOffice = new RadioButton();
        rbMSWord = new RadioButton();
        gbxColorScheme = new GroupBox();
        rbtnPale = new RadioButton();
        rbtnDark = new RadioButton();
        rbtnBlue = new RadioButton();
        rbtnGrey = new RadioButton();
        tabControl = new TabControl();
        tpAnrufMon = new TabPage();
        lblFRITZBoxMonitor = new Label();
        gbxIPAddress = new GroupBox();
        labelCommaSep = new Label();
        labelMSNs = new Label();
        tbCalledNumbers = new TextBox();
        ckbFritzPlaySound = new CheckBox();
        ckbMonitorContactsFirst = new CheckBox();
        lblFritzBoxHost = new Label();
        ckbFritzMonitorEnabled = new CheckBox();
        iPv4AddressControl = new Adressen.cls.IPv4AddressControl();
        tpHotkey = new TabPage();
        gbHotkey = new GroupBox();
        ckbGlobalHotkey = new CheckBox();
        lblKeyPrefix = new Label();
        cbxHotkeyKey = new ComboBox();
        lblInfo = new Label();
        tpSicherung.SuspendLayout();
        gbBackupZip.SuspendLayout();
        gbBackupDaily.SuspendLayout();
        tpAskBefore.SuspendLayout();
        gbxMin2Tray.SuspendLayout();
        gbxAskEnvelope.SuspendLayout();
        gbxAskLocal.SuspendLayout();
        tpWatchFolder.SuspendLayout();
        gbWatchFolder.SuspendLayout();
        tpAutostart.SuspendLayout();
        gbBirthdayRemind.SuspendLayout();
        groupBox1.SuspendLayout();
        gbxContactsAutoload.SuspendLayout();
        tpAdressen.SuspendLayout();
        groupBox.SuspendLayout();
        gbDatabaseFolder.SuspendLayout();
        tpAllgemein.SuspendLayout();
        gbxFontSize.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)nudFontSize).BeginInit();
        gbTextProcessing.SuspendLayout();
        gbxColorScheme.SuspendLayout();
        tabControl.SuspendLayout();
        tpAnrufMon.SuspendLayout();
        gbxIPAddress.SuspendLayout();
        tpHotkey.SuspendLayout();
        gbHotkey.SuspendLayout();
        SuspendLayout();
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
        // folderBrowserDialog
        // 
        folderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer;
        folderBrowserDialog.UseDescriptionForTitle = true;
        // 
        // openFileDialog
        // 
        openFileDialog.DefaultExt = "adb";
        openFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";
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
        // tpSicherung
        // 
        tpSicherung.BackColor = SystemColors.ControlLightLight;
        tpSicherung.BorderStyle = BorderStyle.FixedSingle;
        tpSicherung.Controls.Add(gbBackupZip);
        tpSicherung.Controls.Add(gbBackupDaily);
        tpSicherung.Location = new Point(114, 4);
        tpSicherung.Name = "tpSicherung";
        tpSicherung.Size = new Size(271, 304);
        tpSicherung.TabIndex = 2;
        tpSicherung.Text = " Sicherung";
        // 
        // gbBackupZip
        // 
        gbBackupZip.Controls.Add(btnZipArchive);
        gbBackupZip.Controls.Add(ckbZipArchive);
        gbBackupZip.Controls.Add(lblZipArchive);
        gbBackupZip.Controls.Add(tbZipArchive);
        gbBackupZip.Location = new Point(6, 191);
        gbBackupZip.Name = "gbBackupZip";
        gbBackupZip.Size = new Size(257, 105);
        gbBackupZip.TabIndex = 18;
        gbBackupZip.TabStop = false;
        gbBackupZip.Text = "Zip-Backup";
        // 
        // btnZipArchive
        // 
        btnZipArchive.Location = new Point(215, 72);
        btnZipArchive.Name = "btnZipArchive";
        btnZipArchive.Size = new Size(36, 25);
        btnZipArchive.TabIndex = 16;
        btnZipArchive.Text = "⚙";
        btnZipArchive.UseVisualStyleBackColor = true;
        btnZipArchive.Click += BtnZipArchive_Click;
        // 
        // ckbZipArchive
        // 
        ckbZipArchive.AutoSize = true;
        ckbZipArchive.Location = new Point(6, 24);
        ckbZipArchive.Name = "ckbZipArchive";
        ckbZipArchive.Size = new Size(233, 23);
        ckbZipArchive.TabIndex = 11;
        ckbZipArchive.Text = "Datenbanken in Zip-Datei sichern";
        ckbZipArchive.UseVisualStyleBackColor = true;
        ckbZipArchive.CheckedChanged += CkbZipArchive_CheckedChanged;
        // 
        // lblZipArchive
        // 
        lblZipArchive.AutoSize = true;
        lblZipArchive.Location = new Point(12, 50);
        lblZipArchive.Name = "lblZipArchive";
        lblZipArchive.Size = new Size(69, 19);
        lblZipArchive.TabIndex = 15;
        lblZipArchive.Text = "Zip-Datei:";
        // 
        // tbZipArchive
        // 
        tbZipArchive.BorderStyle = BorderStyle.FixedSingle;
        tbZipArchive.Location = new Point(6, 72);
        tbZipArchive.Name = "tbZipArchive";
        tbZipArchive.Size = new Size(203, 25);
        tbZipArchive.TabIndex = 9;
        tbZipArchive.TextChanged += TbZipArchive_TextChanged;
        tbZipArchive.Validating += TbZipArchive_Validating;
        // 
        // gbBackupDaily
        // 
        gbBackupDaily.Controls.Add(tbBackupFolder);
        gbBackupDaily.Controls.Add(lblBackupFolder);
        gbBackupDaily.Controls.Add(btnExplorer);
        gbBackupDaily.Controls.Add(lblBackup);
        gbBackupDaily.Controls.Add(ckbBackup);
        gbBackupDaily.Controls.Add(btnBackupFolder);
        gbBackupDaily.Location = new Point(6, 6);
        gbBackupDaily.Name = "gbBackupDaily";
        gbBackupDaily.Size = new Size(257, 176);
        gbBackupDaily.TabIndex = 17;
        gbBackupDaily.TabStop = false;
        gbBackupDaily.Text = "Tagessicherung";
        // 
        // tbBackupFolder
        // 
        tbBackupFolder.BorderStyle = BorderStyle.FixedSingle;
        tbBackupFolder.Location = new Point(6, 72);
        tbBackupFolder.Name = "tbBackupFolder";
        tbBackupFolder.Size = new Size(203, 25);
        tbBackupFolder.TabIndex = 0;
        tbBackupFolder.TextChanged += TbBackupFolder_TextChanged;
        // 
        // lblBackupFolder
        // 
        lblBackupFolder.AutoSize = true;
        lblBackupFolder.Location = new Point(6, 50);
        lblBackupFolder.Name = "lblBackupFolder";
        lblBackupFolder.Size = new Size(119, 19);
        lblBackupFolder.TabIndex = 8;
        lblBackupFolder.Text = "Sicherungsordner:";
        // 
        // btnExplorer
        // 
        btnExplorer.Location = new Point(6, 144);
        btnExplorer.Name = "btnExplorer";
        btnExplorer.Size = new Size(245, 26);
        btnExplorer.TabIndex = 7;
        btnExplorer.Text = "Sicherungsordner anzeigen";
        btnExplorer.UseVisualStyleBackColor = true;
        btnExplorer.Click += BtnExplorer_Click;
        // 
        // lblBackup
        // 
        lblBackup.Location = new Point(6, 100);
        lblBackup.Name = "lblBackup";
        lblBackup.Size = new Size(245, 41);
        lblBackup.TabIndex = 6;
        lblBackup.Text = "Das Backup erfolgt in wochentäglichen Unterordnern mit jeweils einer Kopie.";
        // 
        // ckbBackup
        // 
        ckbBackup.AutoSize = true;
        ckbBackup.Location = new Point(6, 24);
        ckbBackup.Name = "ckbBackup";
        ckbBackup.Size = new Size(235, 23);
        ckbBackup.TabIndex = 2;
        ckbBackup.Text = "Daten täglich automatisch sichern";
        ckbBackup.UseVisualStyleBackColor = true;
        ckbBackup.CheckedChanged += CkbBackup_CheckedChanged;
        // 
        // btnBackupFolder
        // 
        btnBackupFolder.Location = new Point(215, 72);
        btnBackupFolder.Name = "btnBackupFolder";
        btnBackupFolder.Size = new Size(36, 25);
        btnBackupFolder.TabIndex = 1;
        btnBackupFolder.Text = "⚙";
        btnBackupFolder.UseVisualStyleBackColor = true;
        btnBackupFolder.Click += BtnBackupFolder_Click;
        // 
        // tpAskBefore
        // 
        tpAskBefore.BackColor = SystemColors.ControlLightLight;
        tpAskBefore.BorderStyle = BorderStyle.FixedSingle;
        tpAskBefore.Controls.Add(gbxMin2Tray);
        tpAskBefore.Controls.Add(gbxAskEnvelope);
        tpAskBefore.Controls.Add(gbxAskLocal);
        tpAskBefore.Location = new Point(114, 4);
        tpAskBefore.Name = "tpAskBefore";
        tpAskBefore.Size = new Size(271, 304);
        tpAskBefore.TabIndex = 5;
        tpAskBefore.Text = " Abfragen";
        // 
        // gbxMin2Tray
        // 
        gbxMin2Tray.Controls.Add(ckbBalloonTipMin2Tray);
        gbxMin2Tray.Location = new Point(6, 176);
        gbxMin2Tray.Name = "gbxMin2Tray";
        gbxMin2Tray.Size = new Size(257, 50);
        gbxMin2Tray.TabIndex = 12;
        gbxMin2Tray.TabStop = false;
        gbxMin2Tray.Text = "Programm ins Tray minimieren";
        // 
        // ckbBalloonTipMin2Tray
        // 
        ckbBalloonTipMin2Tray.AutoSize = true;
        ckbBalloonTipMin2Tray.Checked = true;
        ckbBalloonTipMin2Tray.CheckState = CheckState.Checked;
        ckbBalloonTipMin2Tray.Location = new Point(6, 24);
        ckbBalloonTipMin2Tray.Name = "ckbBalloonTipMin2Tray";
        ckbBalloonTipMin2Tray.Size = new Size(246, 23);
        ckbBalloonTipMin2Tray.TabIndex = 6;
        ckbBalloonTipMin2Tray.Text = "Info zum Wiederherstellen anzeigen";
        ckbBalloonTipMin2Tray.UseVisualStyleBackColor = true;
        // 
        // gbxAskEnvelope
        // 
        gbxAskEnvelope.Controls.Add(ckbAskPrintEnvelope);
        gbxAskEnvelope.Location = new Point(6, 120);
        gbxAskEnvelope.Name = "gbxAskEnvelope";
        gbxAskEnvelope.Size = new Size(257, 50);
        gbxAskEnvelope.TabIndex = 11;
        gbxAskEnvelope.TabStop = false;
        gbxAskEnvelope.Text = "Briefumschläge";
        // 
        // ckbAskPrintEnvelope
        // 
        ckbAskPrintEnvelope.AutoSize = true;
        ckbAskPrintEnvelope.Checked = true;
        ckbAskPrintEnvelope.CheckState = CheckState.Checked;
        ckbAskPrintEnvelope.Location = new Point(6, 24);
        ckbAskPrintEnvelope.Name = "ckbAskPrintEnvelope";
        ckbAskPrintEnvelope.Size = new Size(244, 23);
        ckbAskPrintEnvelope.TabIndex = 6;
        ckbAskPrintEnvelope.Text = "Abfrage vor dem Umschlagdrucken";
        ckbAskPrintEnvelope.UseVisualStyleBackColor = true;
        // 
        // gbxAskLocal
        // 
        gbxAskLocal.Controls.Add(ckbAskBeforeDelete);
        gbxAskLocal.Controls.Add(ckbAskBeforeSaveSQLExpander);
        gbxAskLocal.Controls.Add(ckbAskBeforeSaveSQL);
        gbxAskLocal.Location = new Point(6, 6);
        gbxAskLocal.Name = "gbxAskLocal";
        gbxAskLocal.Size = new Size(257, 108);
        gbxAskLocal.TabIndex = 10;
        gbxAskLocal.TabStop = false;
        gbxAskLocal.Text = "Lokale Adressen";
        // 
        // ckbAskBeforeDelete
        // 
        ckbAskBeforeDelete.AutoSize = true;
        ckbAskBeforeDelete.Checked = true;
        ckbAskBeforeDelete.CheckState = CheckState.Checked;
        ckbAskBeforeDelete.Location = new Point(6, 24);
        ckbAskBeforeDelete.Name = "ckbAskBeforeDelete";
        ckbAskBeforeDelete.Size = new Size(248, 23);
        ckbAskBeforeDelete.TabIndex = 8;
        ckbAskBeforeDelete.Text = "Sicherheitsabfrage vor dem Löschen";
        ckbAskBeforeDelete.UseVisualStyleBackColor = true;
        // 
        // ckbAskBeforeSaveSQLExpander
        // 
        ckbAskBeforeSaveSQLExpander.AutoSize = true;
        ckbAskBeforeSaveSQLExpander.Location = new Point(6, 82);
        ckbAskBeforeSaveSQLExpander.Name = "ckbAskBeforeSaveSQLExpander";
        ckbAskBeforeSaveSQLExpander.Size = new Size(247, 23);
        ckbAskBeforeSaveSQLExpander.TabIndex = 9;
        ckbAskBeforeSaveSQLExpander.Text = "Änderungdetailanzeige ermöglichen";
        ckbAskBeforeSaveSQLExpander.UseVisualStyleBackColor = true;
        // 
        // ckbAskBeforeSaveSQL
        // 
        ckbAskBeforeSaveSQL.AutoSize = true;
        ckbAskBeforeSaveSQL.Location = new Point(6, 53);
        ckbAskBeforeSaveSQL.Name = "ckbAskBeforeSaveSQL";
        ckbAskBeforeSaveSQL.Size = new Size(249, 23);
        ckbAskBeforeSaveSQL.TabIndex = 7;
        ckbAskBeforeSaveSQL.Text = "Abfrage vor Datenbankspeicherung ";
        ckbAskBeforeSaveSQL.UseVisualStyleBackColor = true;
        ckbAskBeforeSaveSQL.CheckedChanged += CkbAskBeforeSaveSQL_CheckedChanged;
        // 
        // tpWatchFolder
        // 
        tpWatchFolder.BackColor = SystemColors.ControlLightLight;
        tpWatchFolder.BorderStyle = BorderStyle.FixedSingle;
        tpWatchFolder.Controls.Add(gbWatchFolder);
        tpWatchFolder.Location = new Point(114, 4);
        tpWatchFolder.Name = "tpWatchFolder";
        tpWatchFolder.Size = new Size(271, 304);
        tpWatchFolder.TabIndex = 4;
        tpWatchFolder.Text = " Dokumente";
        // 
        // gbWatchFolder
        // 
        gbWatchFolder.Controls.Add(ckbWatchFolder);
        gbWatchFolder.Controls.Add(lblWatcherInfo);
        gbWatchFolder.Controls.Add(tbWatchFolder);
        gbWatchFolder.Controls.Add(btnWatchFolder);
        gbWatchFolder.Location = new Point(6, 6);
        gbWatchFolder.Name = "gbWatchFolder";
        gbWatchFolder.Size = new Size(257, 292);
        gbWatchFolder.TabIndex = 14;
        gbWatchFolder.TabStop = false;
        gbWatchFolder.Text = "Dokumentenordner";
        // 
        // ckbWatchFolder
        // 
        ckbWatchFolder.AutoSize = true;
        ckbWatchFolder.Location = new Point(6, 24);
        ckbWatchFolder.Name = "ckbWatchFolder";
        ckbWatchFolder.Size = new Size(225, 23);
        ckbWatchFolder.TabIndex = 11;
        ckbWatchFolder.Text = "Auf Veränderungen überwachen";
        ckbWatchFolder.UseVisualStyleBackColor = true;
        ckbWatchFolder.CheckedChanged += CkbWatchFolder_CheckedChanged;
        // 
        // lblWatcherInfo
        // 
        lblWatcherInfo.Location = new Point(3, 79);
        lblWatcherInfo.Name = "lblWatcherInfo";
        lblWatcherInfo.Size = new Size(250, 211);
        lblWatcherInfo.TabIndex = 13;
        lblWatcherInfo.Text = resources.GetString("lblWatcherInfo.Text");
        // 
        // tbWatchFolder
        // 
        tbWatchFolder.BorderStyle = BorderStyle.FixedSingle;
        tbWatchFolder.Location = new Point(6, 53);
        tbWatchFolder.Name = "tbWatchFolder";
        tbWatchFolder.Size = new Size(210, 25);
        tbWatchFolder.TabIndex = 9;
        // 
        // btnWatchFolder
        // 
        btnWatchFolder.Location = new Point(215, 53);
        btnWatchFolder.Name = "btnWatchFolder";
        btnWatchFolder.Size = new Size(36, 25);
        btnWatchFolder.TabIndex = 10;
        btnWatchFolder.Text = "⚙";
        btnWatchFolder.UseVisualStyleBackColor = true;
        btnWatchFolder.Click += BtnWatchFolder_Click;
        // 
        // tpAutostart
        // 
        tpAutostart.BackColor = SystemColors.ControlLightLight;
        tpAutostart.BorderStyle = BorderStyle.FixedSingle;
        tpAutostart.Controls.Add(gbBirthdayRemind);
        tpAutostart.Controls.Add(groupBox1);
        tpAutostart.Controls.Add(gbxContactsAutoload);
        tpAutostart.Location = new Point(114, 4);
        tpAutostart.Name = "tpAutostart";
        tpAutostart.Padding = new Padding(3);
        tpAutostart.Size = new Size(271, 304);
        tpAutostart.TabIndex = 1;
        tpAutostart.Text = " Autostart";
        // 
        // gbBirthdayRemind
        // 
        gbBirthdayRemind.Controls.Add(lblBirtdayRemind);
        gbBirthdayRemind.Controls.Add(ckbBirthdayRemind);
        gbBirthdayRemind.Location = new Point(6, 224);
        gbBirthdayRemind.Name = "gbBirthdayRemind";
        gbBirthdayRemind.Size = new Size(257, 72);
        gbBirthdayRemind.TabIndex = 10;
        gbBirthdayRemind.TabStop = false;
        gbBirthdayRemind.Text = "Geburtstagserinnerung";
        // 
        // lblBirtdayRemind
        // 
        lblBirtdayRemind.AutoSize = true;
        lblBirtdayRemind.Location = new Point(3, 47);
        lblBirtdayRemind.Name = "lblBirtdayRemind";
        lblBirtdayRemind.Size = new Size(251, 19);
        lblBirtdayRemind.TabIndex = 11;
        lblBirtdayRemind.Text = "Option greift nur, wenn Reminder aktiv.";
        // 
        // ckbBirthdayRemind
        // 
        ckbBirthdayRemind.AutoSize = true;
        ckbBirthdayRemind.Location = new Point(6, 24);
        ckbBirthdayRemind.Name = "ckbBirthdayRemind";
        ckbBirthdayRemind.Size = new Size(240, 23);
        ckbBirthdayRemind.TabIndex = 2;
        ckbBirthdayRemind.Text = "Liste mind. einmal täglich anzeigen";
        ckbBirthdayRemind.UseVisualStyleBackColor = true;
        // 
        // groupBox1
        // 
        groupBox1.Controls.Add(ckbMin2Tray);
        groupBox1.Controls.Add(ckbAutostart);
        groupBox1.Controls.Add(lblAutostart);
        groupBox1.Location = new Point(6, 7);
        groupBox1.Name = "groupBox1";
        groupBox1.Size = new Size(257, 114);
        groupBox1.TabIndex = 7;
        groupBox1.TabStop = false;
        groupBox1.Text = "Adressen && Kontakte (das Programm)";
        // 
        // ckbMin2Tray
        // 
        ckbMin2Tray.AutoSize = true;
        ckbMin2Tray.Location = new Point(6, 48);
        ckbMin2Tray.Name = "ckbMin2Tray";
        ckbMin2Tray.Size = new Size(233, 23);
        ckbMin2Tray.TabIndex = 3;
        ckbMin2Tray.Text = "und sofort in das Tray minimieren";
        ckbMin2Tray.UseVisualStyleBackColor = true;
        ckbMin2Tray.CheckedChanged += CkbMin2Tray_CheckedChanged;
        // 
        // ckbAutostart
        // 
        ckbAutostart.AutoSize = true;
        ckbAutostart.Location = new Point(6, 24);
        ckbAutostart.Name = "ckbAutostart";
        ckbAutostart.Size = new Size(238, 23);
        ckbAutostart.TabIndex = 2;
        ckbAutostart.Text = "Bei Benutzeranmeldung ausführen";
        ckbAutostart.UseVisualStyleBackColor = true;
        ckbAutostart.CheckedChanged += CkbAutostart_CheckedChanged;
        // 
        // lblAutostart
        // 
        lblAutostart.AutoSize = true;
        lblAutostart.Location = new Point(3, 71);
        lblAutostart.Name = "lblAutostart";
        lblAutostart.Size = new Size(243, 38);
        lblAutostart.TabIndex = 8;
        lblAutostart.Text = "Sinnvoll, wenn Anrufmonitor oder Do-\r\nkumentenordnerüberwachung aktiv.";
        // 
        // gbxContactsAutoload
        // 
        gbxContactsAutoload.Controls.Add(ckbContactsAutoload);
        gbxContactsAutoload.Controls.Add(labelAutoAdressen);
        gbxContactsAutoload.Location = new Point(6, 127);
        gbxContactsAutoload.Name = "gbxContactsAutoload";
        gbxContactsAutoload.Size = new Size(257, 91);
        gbxContactsAutoload.TabIndex = 6;
        gbxContactsAutoload.TabStop = false;
        gbxContactsAutoload.Text = "Google Kontakte";
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
        // labelAutoAdressen
        // 
        labelAutoAdressen.AutoSize = true;
        labelAutoAdressen.Location = new Point(3, 47);
        labelAutoAdressen.Name = "labelAutoAdressen";
        labelAutoAdressen.Size = new Size(252, 38);
        labelAutoAdressen.TabIndex = 9;
        labelAutoAdressen.Text = "Welche lokale Adressdatei automatisch\r\ngeladen wird, ergibt sich bei „Adressen“.";
        // 
        // tpAdressen
        // 
        tpAdressen.BackColor = SystemColors.ControlLightLight;
        tpAdressen.BorderStyle = BorderStyle.FixedSingle;
        tpAdressen.Controls.Add(lblToggleDatabase);
        tpAdressen.Controls.Add(groupBox);
        tpAdressen.Controls.Add(gbDatabaseFolder);
        tpAdressen.Location = new Point(114, 4);
        tpAdressen.Name = "tpAdressen";
        tpAdressen.Padding = new Padding(3);
        tpAdressen.Size = new Size(271, 304);
        tpAdressen.TabIndex = 0;
        tpAdressen.Text = " Adressen";
        // 
        // lblToggleDatabase
        // 
        lblToggleDatabase.Location = new Point(12, 227);
        lblToggleDatabase.Name = "lblToggleDatabase";
        lblToggleDatabase.Size = new Size(245, 63);
        lblToggleDatabase.TabIndex = 3;
        lblToggleDatabase.Text = "Mit der F12-Taste lässt sich zwischen\r\nzwei lokalen Datenbanken wechseln.\r\nSiehe Menü Datei > Zuletzt geöffnet";
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
        tbStandard.Validating += TbStandard_Validating;
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
        // tpAllgemein
        // 
        tpAllgemein.BackColor = SystemColors.ControlLightLight;
        tpAllgemein.BorderStyle = BorderStyle.FixedSingle;
        tpAllgemein.Controls.Add(gbxFontSize);
        tpAllgemein.Controls.Add(gbTextProcessing);
        tpAllgemein.Controls.Add(gbxColorScheme);
        tpAllgemein.Location = new Point(114, 4);
        tpAllgemein.Name = "tpAllgemein";
        tpAllgemein.Size = new Size(271, 304);
        tpAllgemein.TabIndex = 3;
        tpAllgemein.Text = " Allgemein";
        // 
        // gbxFontSize
        // 
        gbxFontSize.Controls.Add(ckbPlaceholderText);
        gbxFontSize.Controls.Add(btnFontReset);
        gbxFontSize.Controls.Add(nudFontSize);
        gbxFontSize.Controls.Add(cbxFontName);
        gbxFontSize.Location = new Point(3, 67);
        gbxFontSize.Name = "gbxFontSize";
        gbxFontSize.Size = new Size(263, 118);
        gbxFontSize.TabIndex = 7;
        gbxFontSize.TabStop = false;
        gbxFontSize.Text = "Schriftart für Textfelder";
        // 
        // ckbPlaceholderText
        // 
        ckbPlaceholderText.AutoSize = true;
        ckbPlaceholderText.Checked = true;
        ckbPlaceholderText.CheckState = CheckState.Checked;
        ckbPlaceholderText.Location = new Point(8, 89);
        ckbPlaceholderText.Name = "ckbPlaceholderText";
        ckbPlaceholderText.Size = new Size(245, 23);
        ckbPlaceholderText.TabIndex = 6;
        ckbPlaceholderText.Text = "Hinweise in leeren Feldern anzeigen";
        ckbPlaceholderText.UseVisualStyleBackColor = true;
        // 
        // btnFontReset
        // 
        btnFontReset.Location = new Point(6, 56);
        btnFontReset.Name = "btnFontReset";
        btnFontReset.Size = new Size(251, 27);
        btnFontReset.TabIndex = 5;
        btnFontReset.Text = "Standard: Segoe UI, Schriftgröße 10";
        btnFontReset.UseVisualStyleBackColor = true;
        btnFontReset.Click += BtnFontReset_Click;
        // 
        // nudFontSize
        // 
        nudFontSize.Location = new Point(202, 25);
        nudFontSize.Maximum = new decimal(new int[] { 12, 0, 0, 0 });
        nudFontSize.Minimum = new decimal(new int[] { 9, 0, 0, 0 });
        nudFontSize.Name = "nudFontSize";
        nudFontSize.Size = new Size(55, 25);
        nudFontSize.TabIndex = 1;
        nudFontSize.TextAlign = HorizontalAlignment.Center;
        nudFontSize.Value = new decimal(new int[] { 10, 0, 0, 0 });
        nudFontSize.ValueChanged += NudFontSize_ValueChanged;
        // 
        // cbxFontName
        // 
        cbxFontName.DrawMode = DrawMode.OwnerDrawFixed;
        cbxFontName.DropDownStyle = ComboBoxStyle.DropDownList;
        cbxFontName.FormattingEnabled = true;
        cbxFontName.ItemHeight = 20;
        cbxFontName.Location = new Point(6, 24);
        cbxFontName.Name = "cbxFontName";
        cbxFontName.Size = new Size(190, 26);
        cbxFontName.TabIndex = 0;
        cbxFontName.DrawItem += CbxFontName_DrawItem;
        cbxFontName.SelectedIndexChanged += CbxFontName_SelectedIndexChanged;
        // 
        // gbTextProcessing
        // 
        gbTextProcessing.Controls.Add(rbManualSelect);
        gbTextProcessing.Controls.Add(rbLibreOffice);
        gbTextProcessing.Controls.Add(rbMSWord);
        gbTextProcessing.Location = new Point(3, 191);
        gbTextProcessing.Name = "gbTextProcessing";
        gbTextProcessing.Size = new Size(263, 108);
        gbTextProcessing.TabIndex = 6;
        gbTextProcessing.TabStop = false;
        gbTextProcessing.Text = "Textverarbeitungsprogramm";
        // 
        // rbManualSelect
        // 
        rbManualSelect.AutoSize = true;
        rbManualSelect.Location = new Point(30, 76);
        rbManualSelect.Name = "rbManualSelect";
        rbManualSelect.Size = new Size(150, 23);
        rbManualSelect.TabIndex = 2;
        rbManualSelect.Text = "Jedesmal auswählen";
        rbManualSelect.UseVisualStyleBackColor = true;
        // 
        // rbLibreOffice
        // 
        rbLibreOffice.AutoSize = true;
        rbLibreOffice.Location = new Point(30, 52);
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
        // tabControl
        // 
        tabControl.Alignment = TabAlignment.Left;
        tabControl.Controls.Add(tpAllgemein);
        tabControl.Controls.Add(tpAdressen);
        tabControl.Controls.Add(tpAutostart);
        tabControl.Controls.Add(tpWatchFolder);
        tabControl.Controls.Add(tpAnrufMon);
        tabControl.Controls.Add(tpHotkey);
        tabControl.Controls.Add(tpAskBefore);
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
        // tpAnrufMon
        // 
        tpAnrufMon.BackColor = SystemColors.ControlLightLight;
        tpAnrufMon.BorderStyle = BorderStyle.FixedSingle;
        tpAnrufMon.Controls.Add(lblFRITZBoxMonitor);
        tpAnrufMon.Controls.Add(gbxIPAddress);
        tpAnrufMon.Location = new Point(114, 4);
        tpAnrufMon.Name = "tpAnrufMon";
        tpAnrufMon.Size = new Size(271, 304);
        tpAnrufMon.TabIndex = 6;
        tpAnrufMon.Text = " Anrufmonitor";
        // 
        // lblFRITZBoxMonitor
        // 
        lblFRITZBoxMonitor.AutoSize = true;
        lblFRITZBoxMonitor.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 0);
        lblFRITZBoxMonitor.Location = new Point(3, 4);
        lblFRITZBoxMonitor.Name = "lblFRITZBoxMonitor";
        lblFRITZBoxMonitor.Size = new Size(249, 114);
        lblFRITZBoxMonitor.TabIndex = 1;
        lblFRITZBoxMonitor.Text = resources.GetString("lblFRITZBoxMonitor.Text");
        // 
        // gbxIPAddress
        // 
        gbxIPAddress.Controls.Add(labelCommaSep);
        gbxIPAddress.Controls.Add(labelMSNs);
        gbxIPAddress.Controls.Add(tbCalledNumbers);
        gbxIPAddress.Controls.Add(ckbFritzPlaySound);
        gbxIPAddress.Controls.Add(ckbMonitorContactsFirst);
        gbxIPAddress.Controls.Add(lblFritzBoxHost);
        gbxIPAddress.Controls.Add(ckbFritzMonitorEnabled);
        gbxIPAddress.Controls.Add(iPv4AddressControl);
        gbxIPAddress.Location = new Point(3, 120);
        gbxIPAddress.Name = "gbxIPAddress";
        gbxIPAddress.Size = new Size(263, 179);
        gbxIPAddress.TabIndex = 0;
        gbxIPAddress.TabStop = false;
        // 
        // labelCommaSep
        // 
        labelCommaSep.AutoSize = true;
        labelCommaSep.Location = new Point(170, 99);
        labelCommaSep.Name = "labelCommaSep";
        labelCommaSep.Size = new Size(87, 19);
        labelCommaSep.TabIndex = 18;
        labelCommaSep.Text = "(kommasep.)";
        // 
        // labelMSNs
        // 
        labelMSNs.AutoSize = true;
        labelMSNs.Location = new Point(3, 75);
        labelMSNs.Name = "labelMSNs";
        labelMSNs.Size = new Size(250, 19);
        labelMSNs.TabIndex = 17;
        labelMSNs.Text = "Eigene Rufnummern (leeres Feld = alle):";
        // 
        // tbCalledNumbers
        // 
        tbCalledNumbers.BorderStyle = BorderStyle.FixedSingle;
        tbCalledNumbers.Location = new Point(6, 97);
        tbCalledNumbers.Name = "tbCalledNumbers";
        tbCalledNumbers.Size = new Size(164, 25);
        tbCalledNumbers.TabIndex = 16;
        // 
        // ckbFritzPlaySound
        // 
        ckbFritzPlaySound.AutoSize = true;
        ckbFritzPlaySound.Location = new Point(6, 128);
        ckbFritzPlaySound.Name = "ckbFritzPlaySound";
        ckbFritzPlaySound.Size = new Size(247, 23);
        ckbFritzPlaySound.TabIndex = 15;
        ckbFritzPlaySound.Text = "Bei Anruf eine Sounddatei abspielen";
        ckbFritzPlaySound.UseVisualStyleBackColor = true;
        // 
        // ckbMonitorContactsFirst
        // 
        ckbMonitorContactsFirst.AutoSize = true;
        ckbMonitorContactsFirst.Location = new Point(6, 154);
        ckbMonitorContactsFirst.Name = "ckbMonitorContactsFirst";
        ckbMonitorContactsFirst.Size = new Size(255, 23);
        ckbMonitorContactsFirst.TabIndex = 14;
        ckbMonitorContactsFirst.Text = "Erst Kontakte, dann Adressen suchen";
        ckbMonitorContactsFirst.UseVisualStyleBackColor = true;
        // 
        // lblFritzBoxHost
        // 
        lblFritzBoxHost.AutoSize = true;
        lblFritzBoxHost.Location = new Point(3, 46);
        lblFritzBoxHost.Name = "lblFritzBoxHost";
        lblFritzBoxHost.Size = new Size(78, 19);
        lblFritzBoxHost.TabIndex = 13;
        lblFritzBoxHost.Text = "IP-Adresse:";
        // 
        // ckbFritzMonitorEnabled
        // 
        ckbFritzMonitorEnabled.AutoSize = true;
        ckbFritzMonitorEnabled.Location = new Point(6, 15);
        ckbFritzMonitorEnabled.Name = "ckbFritzMonitorEnabled";
        ckbFritzMonitorEnabled.Size = new Size(247, 23);
        ckbFritzMonitorEnabled.TabIndex = 12;
        ckbFritzMonitorEnabled.Text = "Auf eingehende Anrufe überwachen";
        ckbFritzMonitorEnabled.UseVisualStyleBackColor = true;
        ckbFritzMonitorEnabled.CheckedChanged += CkbFritzMonitorEnabled_CheckedChanged;
        // 
        // iPv4AddressControl
        // 
        iPv4AddressControl.BackColor = SystemColors.Window;
        iPv4AddressControl.BorderStyle = BorderStyle.FixedSingle;
        iPv4AddressControl.Font = new Font("Consolas", 10F);
        iPv4AddressControl.Location = new Point(90, 44);
        iPv4AddressControl.Name = "iPv4AddressControl";
        iPv4AddressControl.Size = new Size(160, 23);
        iPv4AddressControl.TabIndex = 0;
        // 
        // tpHotkey
        // 
        tpHotkey.BackColor = SystemColors.ControlLightLight;
        tpHotkey.BorderStyle = BorderStyle.FixedSingle;
        tpHotkey.Controls.Add(gbHotkey);
        tpHotkey.Controls.Add(lblInfo);
        tpHotkey.Location = new Point(114, 4);
        tpHotkey.Name = "tpHotkey";
        tpHotkey.Size = new Size(271, 304);
        tpHotkey.TabIndex = 7;
        tpHotkey.Text = " Hotkey";
        // 
        // gbHotkey
        // 
        gbHotkey.Controls.Add(ckbGlobalHotkey);
        gbHotkey.Controls.Add(lblKeyPrefix);
        gbHotkey.Controls.Add(cbxHotkeyKey);
        gbHotkey.Location = new Point(6, 140);
        gbHotkey.Name = "gbHotkey";
        gbHotkey.Size = new Size(257, 79);
        gbHotkey.TabIndex = 4;
        gbHotkey.TabStop = false;
        // 
        // ckbGlobalHotkey
        // 
        ckbGlobalHotkey.AutoSize = true;
        ckbGlobalHotkey.Location = new Point(6, 17);
        ckbGlobalHotkey.Name = "ckbGlobalHotkey";
        ckbGlobalHotkey.Size = new Size(193, 23);
        ckbGlobalHotkey.TabIndex = 3;
        ckbGlobalHotkey.Text = "Globalen Hotkey aktivieren";
        ckbGlobalHotkey.UseVisualStyleBackColor = true;
        ckbGlobalHotkey.CheckedChanged += CkbGlobalHotkey_CheckedChanged;
        // 
        // lblKeyPrefix
        // 
        lblKeyPrefix.AutoSize = true;
        lblKeyPrefix.Location = new Point(2, 49);
        lblKeyPrefix.Name = "lblKeyPrefix";
        lblKeyPrefix.Size = new Size(193, 19);
        lblKeyPrefix.TabIndex = 1;
        lblKeyPrefix.Text = "Tastenkombination: Strg+Alt+";
        // 
        // cbxHotkeyKey
        // 
        cbxHotkeyKey.DropDownStyle = ComboBoxStyle.DropDownList;
        cbxHotkeyKey.FormattingEnabled = true;
        cbxHotkeyKey.Items.AddRange(new object[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Z" });
        cbxHotkeyKey.Location = new Point(195, 46);
        cbxHotkeyKey.Name = "cbxHotkeyKey";
        cbxHotkeyKey.Size = new Size(56, 25);
        cbxHotkeyKey.TabIndex = 2;
        // 
        // lblInfo
        // 
        lblInfo.AutoSize = true;
        lblInfo.Location = new Point(3, 4);
        lblInfo.Name = "lblInfo";
        lblInfo.Size = new Size(266, 133);
        lblInfo.TabIndex = 0;
        lblInfo.Text = resources.GetString("lblInfo.Text");
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
        tpSicherung.ResumeLayout(false);
        gbBackupZip.ResumeLayout(false);
        gbBackupZip.PerformLayout();
        gbBackupDaily.ResumeLayout(false);
        gbBackupDaily.PerformLayout();
        tpAskBefore.ResumeLayout(false);
        gbxMin2Tray.ResumeLayout(false);
        gbxMin2Tray.PerformLayout();
        gbxAskEnvelope.ResumeLayout(false);
        gbxAskEnvelope.PerformLayout();
        gbxAskLocal.ResumeLayout(false);
        gbxAskLocal.PerformLayout();
        tpWatchFolder.ResumeLayout(false);
        gbWatchFolder.ResumeLayout(false);
        gbWatchFolder.PerformLayout();
        tpAutostart.ResumeLayout(false);
        gbBirthdayRemind.ResumeLayout(false);
        gbBirthdayRemind.PerformLayout();
        groupBox1.ResumeLayout(false);
        groupBox1.PerformLayout();
        gbxContactsAutoload.ResumeLayout(false);
        gbxContactsAutoload.PerformLayout();
        tpAdressen.ResumeLayout(false);
        groupBox.ResumeLayout(false);
        groupBox.PerformLayout();
        gbDatabaseFolder.ResumeLayout(false);
        gbDatabaseFolder.PerformLayout();
        tpAllgemein.ResumeLayout(false);
        gbxFontSize.ResumeLayout(false);
        gbxFontSize.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)nudFontSize).EndInit();
        gbTextProcessing.ResumeLayout(false);
        gbTextProcessing.PerformLayout();
        gbxColorScheme.ResumeLayout(false);
        gbxColorScheme.PerformLayout();
        tabControl.ResumeLayout(false);
        tpAnrufMon.ResumeLayout(false);
        tpAnrufMon.PerformLayout();
        gbxIPAddress.ResumeLayout(false);
        gbxIPAddress.PerformLayout();
        tpHotkey.ResumeLayout(false);
        tpHotkey.PerformLayout();
        gbHotkey.ResumeLayout(false);
        gbHotkey.PerformLayout();
        ResumeLayout(false);
    }

    #endregion
    private Button btnCancel;
    private FolderBrowserDialog folderBrowserDialog;
    private OpenFileDialog openFileDialog;
    private Button btnOK;
    private TabPage tpSicherung;
    private Button btnZipArchive;
    private Label lblZipArchive;
    private CheckBox ckbZipArchive;
    private TextBox tbZipArchive;
    private TextBox tbBackupFolder;
    private Label lblBackupFolder;
    private Button btnExplorer;
    private Label lblBackup;
    private CheckBox ckbBackup;
    private Button btnBackupFolder;
    private TabPage tpAskBefore;
    private CheckBox ckbAskBeforeSaveSQLExpander;
    private CheckBox ckbAskBeforeDelete;
    private CheckBox ckbAskBeforeSaveSQL;
    private CheckBox ckbAskPrintEnvelope;
    private TabPage tpWatchFolder;
    private Label lblWatcherInfo;
    private CheckBox ckbWatchFolder;
    private Button btnWatchFolder;
    private TextBox tbWatchFolder;
    private TabPage tpAutostart;
    private GroupBox gbxContactsAutoload;
    private CheckBox ckbContactsAutoload;
    private TabPage tpAdressen;
    private Label lblToggleDatabase;
    private GroupBox groupBox;
    private Button btnStandardFile;
    private TextBox tbStandard;
    private RadioButton rbStandard;
    private RadioButton rbRecent;
    private RadioButton rbEmpty;
    private GroupBox gbDatabaseFolder;
    private Button btnDatabaseFolder;
    private TextBox tbDatabaseFolder;
    private TabPage tpAllgemein;
    private GroupBox gbxFontSize;
    private Button btnFontReset;
    private NumericUpDown nudFontSize;
    private ComboBox cbxFontName;
    private GroupBox gbTextProcessing;
    private RadioButton rbManualSelect;
    private RadioButton rbLibreOffice;
    private RadioButton rbMSWord;
    private GroupBox gbxColorScheme;
    private RadioButton rbtnPale;
    private RadioButton rbtnDark;
    private RadioButton rbtnBlue;
    private RadioButton rbtnGrey;
    private TabControl tabControl;
    private GroupBox gbxAskLocal;
    private GroupBox gbxAskEnvelope;
    private CheckBox ckbPlaceholderText;
    private TabPage tpAnrufMon;
    private GroupBox gbxIPAddress;
    private cls.IPv4AddressControl iPv4AddressControl;
    private CheckBox ckbFritzMonitorEnabled;
    private Label lblFritzBoxHost;
    private Label lblFRITZBoxMonitor;
    private CheckBox ckbMonitorContactsFirst;
    private CheckBox ckbFritzPlaySound;
    private GroupBox groupBox1;
    private CheckBox ckbAutostart;
    private CheckBox ckbMin2Tray;
    private Label lblAutostart;
    private Label labelAutoAdressen;
    private GroupBox gbxMin2Tray;
    private CheckBox ckbBalloonTipMin2Tray;
    private Label labelMSNs;
    private TextBox tbCalledNumbers;
    private Label labelCommaSep;
    private TabPage tpHotkey;
    private Label lblInfo;
    private CheckBox ckbGlobalHotkey;
    private ComboBox cbxHotkeyKey;
    private Label lblKeyPrefix;
    private GroupBox gbHotkey;
    private GroupBox gbBirthdayRemind;
    private CheckBox ckbBirthdayRemind;
    private Label lblBirtdayRemind;
    private GroupBox gbWatchFolder;
    private GroupBox gbBackupZip;
    private GroupBox gbBackupDaily;
}