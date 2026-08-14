namespace Adressen.frm;

partial class FrmAdvSearch
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
        groupBox1 = new GroupBox();
        btnResetDate = new Button();
        label11 = new Label();
        yearSlider2 = new Adressen.cls.YearSlider();
        yearSlider1 = new Adressen.cls.YearSlider();
        tbPLZbis = new TextBox();
        labelBis = new Label();
        tbPLZvon = new TextBox();
        label9 = new Label();
        tbOrt = new TextBox();
        label8 = new Label();
        tbStrasse = new TextBox();
        label7 = new Label();
        tbUnternehmen = new TextBox();
        label6 = new Label();
        tbAnrede = new TextBox();
        label5 = new Label();
        tbTitel = new TextBox();
        label4 = new Label();
        tbNickname = new TextBox();
        label3 = new Label();
        tbNachname = new TextBox();
        label2 = new Label();
        tbVorname = new TextBox();
        label1 = new Label();
        groupBox2 = new GroupBox();
        rbExact = new RadioButton();
        rbStartwith = new RadioButton();
        rbContains = new RadioButton();
        groupBox3 = new GroupBox();
        rbOR = new RadioButton();
        rbAND = new RadioButton();
        btnSearch = new Button();
        btnReset = new Button();
        btnCancel = new Button();
        panel = new Panel();
        cbRefineSearch = new CheckBox();
        groupBox1.SuspendLayout();
        groupBox2.SuspendLayout();
        groupBox3.SuspendLayout();
        SuspendLayout();
        // 
        // groupBox1
        // 
        groupBox1.BackColor = SystemColors.Window;
        groupBox1.Controls.Add(btnResetDate);
        groupBox1.Controls.Add(label11);
        groupBox1.Controls.Add(yearSlider2);
        groupBox1.Controls.Add(yearSlider1);
        groupBox1.Controls.Add(tbPLZbis);
        groupBox1.Controls.Add(labelBis);
        groupBox1.Controls.Add(tbPLZvon);
        groupBox1.Controls.Add(label9);
        groupBox1.Controls.Add(tbOrt);
        groupBox1.Controls.Add(label8);
        groupBox1.Controls.Add(tbStrasse);
        groupBox1.Controls.Add(label7);
        groupBox1.Controls.Add(tbUnternehmen);
        groupBox1.Controls.Add(label6);
        groupBox1.Controls.Add(tbAnrede);
        groupBox1.Controls.Add(label5);
        groupBox1.Controls.Add(tbTitel);
        groupBox1.Controls.Add(label4);
        groupBox1.Controls.Add(tbNickname);
        groupBox1.Controls.Add(label3);
        groupBox1.Controls.Add(tbNachname);
        groupBox1.Controls.Add(label2);
        groupBox1.Controls.Add(tbVorname);
        groupBox1.Controls.Add(label1);
        groupBox1.Location = new Point(12, 12);
        groupBox1.Name = "groupBox1";
        groupBox1.Size = new Size(369, 336);
        groupBox1.TabIndex = 0;
        groupBox1.TabStop = false;
        groupBox1.Text = "Suchfelder";
        // 
        // btnResetDate
        // 
        btnResetDate.Enabled = false;
        btnResetDate.Location = new Point(216, 302);
        btnResetDate.Margin = new Padding(3, 2, 3, 1);
        btnResetDate.Name = "btnResetDate";
        btnResetDate.Size = new Size(26, 26);
        btnResetDate.TabIndex = 24;
        btnResetDate.TabStop = false;
        btnResetDate.UseVisualStyleBackColor = true;
        btnResetDate.Click += BtnResetDate_Click;
        // 
        // label11
        // 
        label11.AutoSize = true;
        label11.Location = new Point(6, 306);
        label11.Name = "label11";
        label11.Size = new Size(84, 19);
        label11.TabIndex = 22;
        label11.Text = "Geburtsjahr:";
        // 
        // yearSlider2
        // 
        yearSlider2.BackColor = SystemColors.Window;
        yearSlider2.BorderStyle = BorderStyle.FixedSingle;
        yearSlider2.Enabled = false;
        yearSlider2.Font = new Font("Segoe UI", 10F);
        yearSlider2.Location = new Point(264, 303);
        yearSlider2.Name = "yearSlider2";
        yearSlider2.Size = new Size(99, 24);
        yearSlider2.TabIndex = 21;
        yearSlider2.Enter += YearText_Enter;
        yearSlider2.Leave += YearText_Leave;
        // 
        // yearSlider1
        // 
        yearSlider1.BackColor = SystemColors.Window;
        yearSlider1.BorderStyle = BorderStyle.FixedSingle;
        yearSlider1.Font = new Font("Segoe UI", 10F);
        yearSlider1.Location = new Point(109, 303);
        yearSlider1.Name = "yearSlider1";
        yearSlider1.Size = new Size(99, 24);
        yearSlider1.TabIndex = 20;
        yearSlider1.RawTextChanged += YearSlider1_RawTextChanged;
        yearSlider1.Enter += YearText_Enter;
        yearSlider1.Leave += YearText_Leave;
        // 
        // tbPLZbis
        // 
        tbPLZbis.BorderStyle = BorderStyle.FixedSingle;
        tbPLZbis.Enabled = false;
        tbPLZbis.Location = new Point(264, 272);
        tbPLZbis.Name = "tbPLZbis";
        tbPLZbis.Size = new Size(99, 25);
        tbPLZbis.TabIndex = 19;
        tbPLZbis.Enter += TextBox_Enter;
        tbPLZbis.Leave += TextBox_Leave;
        // 
        // labelBis
        // 
        labelBis.AutoSize = true;
        labelBis.Enabled = false;
        labelBis.Location = new Point(216, 275);
        labelBis.Name = "labelBis";
        labelBis.Size = new Size(29, 19);
        labelBis.TabIndex = 18;
        labelBis.Text = "bis:";
        // 
        // tbPLZvon
        // 
        tbPLZvon.BorderStyle = BorderStyle.FixedSingle;
        tbPLZvon.Location = new Point(109, 272);
        tbPLZvon.Name = "tbPLZvon";
        tbPLZvon.Size = new Size(99, 25);
        tbPLZvon.TabIndex = 17;
        tbPLZvon.TextChanged += TbPLZvon_TextChanged;
        tbPLZvon.Enter += TextBox_Enter;
        tbPLZvon.Leave += TextBox_Leave;
        // 
        // label9
        // 
        label9.AutoSize = true;
        label9.Location = new Point(6, 275);
        label9.Name = "label9";
        label9.Size = new Size(70, 19);
        label9.TabIndex = 16;
        label9.Text = "PLZ (von):";
        // 
        // tbOrt
        // 
        tbOrt.BorderStyle = BorderStyle.FixedSingle;
        tbOrt.Location = new Point(109, 241);
        tbOrt.Name = "tbOrt";
        tbOrt.Size = new Size(254, 25);
        tbOrt.TabIndex = 15;
        tbOrt.Enter += TextBox_Enter;
        tbOrt.Leave += TextBox_Leave;
        // 
        // label8
        // 
        label8.AutoSize = true;
        label8.Location = new Point(6, 244);
        label8.Name = "label8";
        label8.Size = new Size(33, 19);
        label8.TabIndex = 14;
        label8.Text = "Ort:";
        // 
        // tbStrasse
        // 
        tbStrasse.BorderStyle = BorderStyle.FixedSingle;
        tbStrasse.Location = new Point(109, 210);
        tbStrasse.Name = "tbStrasse";
        tbStrasse.Size = new Size(254, 25);
        tbStrasse.TabIndex = 13;
        tbStrasse.Enter += TextBox_Enter;
        tbStrasse.Leave += TextBox_Leave;
        // 
        // label7
        // 
        label7.AutoSize = true;
        label7.Location = new Point(6, 213);
        label7.Name = "label7";
        label7.Size = new Size(51, 19);
        label7.TabIndex = 12;
        label7.Text = "Adresse:";
        // 
        // tbUnternehmen
        // 
        tbUnternehmen.BorderStyle = BorderStyle.FixedSingle;
        tbUnternehmen.Location = new Point(109, 179);
        tbUnternehmen.Name = "tbUnternehmen";
        tbUnternehmen.Size = new Size(254, 25);
        tbUnternehmen.TabIndex = 11;
        tbUnternehmen.Enter += TextBox_Enter;
        tbUnternehmen.Leave += TextBox_Leave;
        // 
        // label6
        // 
        label6.AutoSize = true;
        label6.Location = new Point(6, 182);
        label6.Name = "label6";
        label6.Size = new Size(97, 19);
        label6.TabIndex = 10;
        label6.Text = "Unternehmen:";
        // 
        // tbAnrede
        // 
        tbAnrede.BorderStyle = BorderStyle.FixedSingle;
        tbAnrede.Location = new Point(109, 148);
        tbAnrede.Name = "tbAnrede";
        tbAnrede.Size = new Size(254, 25);
        tbAnrede.TabIndex = 9;
        tbAnrede.Enter += TextBox_Enter;
        tbAnrede.Leave += TextBox_Leave;
        // 
        // label5
        // 
        label5.AutoSize = true;
        label5.Location = new Point(6, 151);
        label5.Name = "label5";
        label5.Size = new Size(56, 19);
        label5.TabIndex = 8;
        label5.Text = "Anrede:";
        // 
        // tbTitel
        // 
        tbTitel.BorderStyle = BorderStyle.FixedSingle;
        tbTitel.Location = new Point(109, 117);
        tbTitel.Name = "tbTitel";
        tbTitel.Size = new Size(254, 25);
        tbTitel.TabIndex = 7;
        tbTitel.Enter += TextBox_Enter;
        tbTitel.Leave += TextBox_Leave;
        // 
        // label4
        // 
        label4.AutoSize = true;
        label4.Location = new Point(6, 120);
        label4.Name = "label4";
        label4.Size = new Size(37, 19);
        label4.TabIndex = 6;
        label4.Text = "Titel:";
        // 
        // tbNickname
        // 
        tbNickname.BorderStyle = BorderStyle.FixedSingle;
        tbNickname.Location = new Point(109, 86);
        tbNickname.Name = "tbNickname";
        tbNickname.Size = new Size(254, 25);
        tbNickname.TabIndex = 5;
        tbNickname.Enter += TextBox_Enter;
        tbNickname.Leave += TextBox_Leave;
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Location = new Point(6, 89);
        label3.Name = "label3";
        label3.Size = new Size(72, 19);
        label3.TabIndex = 4;
        label3.Text = "Nickname:";
        // 
        // tbNachname
        // 
        tbNachname.BorderStyle = BorderStyle.FixedSingle;
        tbNachname.Location = new Point(109, 55);
        tbNachname.Name = "tbNachname";
        tbNachname.Size = new Size(254, 25);
        tbNachname.TabIndex = 3;
        tbNachname.Enter += TextBox_Enter;
        tbNachname.Leave += TextBox_Leave;
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(6, 58);
        label2.Name = "label2";
        label2.Size = new Size(77, 19);
        label2.TabIndex = 2;
        label2.Text = "Nachname:";
        // 
        // tbVorname
        // 
        tbVorname.BorderStyle = BorderStyle.FixedSingle;
        tbVorname.Location = new Point(109, 24);
        tbVorname.Name = "tbVorname";
        tbVorname.Size = new Size(254, 25);
        tbVorname.TabIndex = 1;
        tbVorname.Leave += TextBox_Leave;
        tbVorname.MouseEnter += TextBox_Enter;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(6, 27);
        label1.Name = "label1";
        label1.Size = new Size(67, 19);
        label1.TabIndex = 0;
        label1.Text = "Vorname:";
        // 
        // groupBox2
        // 
        groupBox2.BackColor = SystemColors.Window;
        groupBox2.Controls.Add(rbExact);
        groupBox2.Controls.Add(rbStartwith);
        groupBox2.Controls.Add(rbContains);
        groupBox2.Location = new Point(12, 354);
        groupBox2.Name = "groupBox2";
        groupBox2.Size = new Size(369, 53);
        groupBox2.TabIndex = 1;
        groupBox2.TabStop = false;
        groupBox2.Text = "Suchmodus";
        // 
        // rbExact
        // 
        rbExact.AutoSize = true;
        rbExact.Location = new Point(216, 24);
        rbExact.Name = "rbExact";
        rbExact.Size = new Size(144, 23);
        rbExact.TabIndex = 2;
        rbExact.Text = "Entspricht genau …";
        rbExact.UseVisualStyleBackColor = true;
        // 
        // rbStartwith
        // 
        rbStartwith.AutoSize = true;
        rbStartwith.Location = new Point(96, 24);
        rbStartwith.Name = "rbStartwith";
        rbStartwith.Size = new Size(112, 23);
        rbStartwith.TabIndex = 1;
        rbStartwith.Text = "Beginnt mit …";
        rbStartwith.UseVisualStyleBackColor = true;
        // 
        // rbContains
        // 
        rbContains.AutoSize = true;
        rbContains.Checked = true;
        rbContains.Location = new Point(6, 24);
        rbContains.Name = "rbContains";
        rbContains.Size = new Size(84, 23);
        rbContains.TabIndex = 0;
        rbContains.TabStop = true;
        rbContains.Text = "Enthält …";
        rbContains.UseVisualStyleBackColor = true;
        // 
        // groupBox3
        // 
        groupBox3.BackColor = SystemColors.Window;
        groupBox3.Controls.Add(rbOR);
        groupBox3.Controls.Add(rbAND);
        groupBox3.Location = new Point(12, 413);
        groupBox3.Name = "groupBox3";
        groupBox3.Size = new Size(369, 53);
        groupBox3.TabIndex = 2;
        groupBox3.TabStop = false;
        groupBox3.Text = "Verknüpfung";
        // 
        // rbOR
        // 
        rbOR.AutoSize = true;
        rbOR.Location = new Point(216, 23);
        rbOR.Name = "rbOR";
        rbOR.Size = new Size(145, 23);
        rbOR.TabIndex = 1;
        rbOR.Text = "Oder (mind. 1 Feld)";
        rbOR.UseVisualStyleBackColor = true;
        // 
        // rbAND
        // 
        rbAND.AutoSize = true;
        rbAND.Checked = true;
        rbAND.Location = new Point(6, 24);
        rbAND.Name = "rbAND";
        rbAND.Size = new Size(204, 23);
        rbAND.TabIndex = 0;
        rbAND.TabStop = true;
        rbAND.Text = "Und (alle ausgefüllten Felder)";
        rbAND.UseVisualStyleBackColor = true;
        // 
        // btnSearch
        // 
        btnSearch.DialogResult = DialogResult.OK;
        btnSearch.Location = new Point(12, 509);
        btnSearch.Name = "btnSearch";
        btnSearch.Size = new Size(137, 27);
        btnSearch.TabIndex = 3;
        btnSearch.Text = "Suche starten";
        btnSearch.UseVisualStyleBackColor = true;
        // 
        // btnReset
        // 
        btnReset.Location = new Point(155, 510);
        btnReset.Name = "btnReset";
        btnReset.Size = new Size(110, 27);
        btnReset.TabIndex = 4;
        btnReset.Text = "Zurücksetzen";
        btnReset.UseVisualStyleBackColor = true;
        btnReset.Click += BtnReset_Click;
        // 
        // btnCancel
        // 
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new Point(271, 510);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(110, 27);
        btnCancel.TabIndex = 5;
        btnCancel.Text = "Abbrechen";
        btnCancel.UseVisualStyleBackColor = true;
        // 
        // panel
        // 
        panel.BackColor = SystemColors.Window;
        panel.Dock = DockStyle.Top;
        panel.Location = new Point(0, 0);
        panel.Name = "panel";
        panel.Size = new Size(393, 474);
        panel.TabIndex = 6;
        // 
        // cbRefineSearch
        // 
        cbRefineSearch.AutoSize = true;
        cbRefineSearch.Enabled = false;
        cbRefineSearch.Location = new Point(18, 480);
        cbRefineSearch.Name = "cbRefineSearch";
        cbRefineSearch.Size = new Size(368, 23);
        cbRefineSearch.TabIndex = 7;
        cbRefineSearch.Text = "Mit vorhandenem Filter bzw. Suchergebnis kombinieren";
        cbRefineSearch.UseVisualStyleBackColor = true;
        // 
        // FrmAdvSearch
        // 
        AcceptButton = btnSearch;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = btnCancel;
        ClientSize = new Size(393, 548);
        Controls.Add(cbRefineSearch);
        Controls.Add(btnCancel);
        Controls.Add(btnReset);
        Controls.Add(btnSearch);
        Controls.Add(groupBox3);
        Controls.Add(groupBox2);
        Controls.Add(groupBox1);
        Controls.Add(panel);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "FrmAdvSearch";
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Erweiterte Suche";
        groupBox1.ResumeLayout(false);
        groupBox1.PerformLayout();
        groupBox2.ResumeLayout(false);
        groupBox2.PerformLayout();
        groupBox3.ResumeLayout(false);
        groupBox3.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private GroupBox groupBox1;
    private TextBox tbNachname;
    private Label label2;
    private TextBox tbVorname;
    private Label label1;
    private TextBox tbStrasse;
    private Label label7;
    private TextBox tbUnternehmen;
    private Label label6;
    private TextBox tbAnrede;
    private Label label5;
    private TextBox tbTitel;
    private Label label4;
    private TextBox tbNickname;
    private Label label3;
    private TextBox tbPLZvon;
    private Label label9;
    private TextBox tbOrt;
    private Label label8;
    private TextBox tbPLZbis;
    private Label labelBis;
    private cls.YearSlider yearSlider1;
    private Label label11;
    private cls.YearSlider yearSlider2;
    private GroupBox groupBox2;
    private RadioButton rbExact;
    private RadioButton rbStartwith;
    private RadioButton rbContains;
    private GroupBox groupBox3;
    private RadioButton rbOR;
    private RadioButton rbAND;
    private Button btnSearch;
    private Button btnReset;
    private Button btnCancel;
    private Button btnResetDate;
    private Panel panel;
    private CheckBox cbRefineSearch;
}