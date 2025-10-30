namespace Adressen.frm;

partial class FrmGroupRename
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
        label = new Label();
        textBox = new TextBox();
        btnOK = new Button();
        btnCancel = new Button();
        SuspendLayout();
        // 
        // label
        // 
        label.Location = new Point(12, 15);
        label.Name = "label";
        label.Size = new Size(98, 25);
        label.TabIndex = 0;
        label.Text = "Neuer Name:";
        // 
        // textBox
        // 
        textBox.Location = new Point(116, 12);
        textBox.Name = "textBox";
        textBox.PlaceholderText = "Gruppenname";
        textBox.Size = new Size(98, 25);
        textBox.TabIndex = 1;
        // 
        // btnOK
        // 
        btnOK.DialogResult = DialogResult.OK;
        btnOK.Location = new Point(12, 43);
        btnOK.Name = "btnOK";
        btnOK.Size = new Size(98, 30);
        btnOK.TabIndex = 2;
        btnOK.Text = "Übernehmen";
        btnOK.UseVisualStyleBackColor = true;
        // 
        // btnCancel
        // 
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new Point(116, 43);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(98, 30);
        btnCancel.TabIndex = 3;
        btnCancel.Text = "Abbrechen";
        btnCancel.UseVisualStyleBackColor = true;
        // 
        // FrmGroupRename
        // 
        AcceptButton = btnOK;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = btnCancel;
        ClientSize = new Size(226, 85);
        Controls.Add(btnCancel);
        Controls.Add(btnOK);
        Controls.Add(textBox);
        Controls.Add(label);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "FrmGroupRename";
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Gruppe umbenennen";
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private Label label;
    private TextBox textBox;
    private Button btnOK;
    private Button btnCancel;
}