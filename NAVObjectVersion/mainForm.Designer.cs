namespace NAVObjectVersion
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.b_PasteClipBoard = new System.Windows.Forms.Button();
            this.chb_UseClipboard = new System.Windows.Forms.CheckBox();
            this.txt_TemplatePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lbl_github = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // b_PasteClipBoard
            // 
            this.b_PasteClipBoard.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.b_PasteClipBoard.Location = new System.Drawing.Point(12, 72);
            this.b_PasteClipBoard.Name = "b_PasteClipBoard";
            this.b_PasteClipBoard.Size = new System.Drawing.Size(305, 153);
            this.b_PasteClipBoard.TabIndex = 0;
            this.b_PasteClipBoard.Text = "Paste NAV Object list";
            this.b_PasteClipBoard.UseVisualStyleBackColor = true;
            this.b_PasteClipBoard.Click += new System.EventHandler(this.b_LoadClipboard_Click);
            // 
            // chb_UseClipboard
            // 
            this.chb_UseClipboard.AutoSize = true;
            this.chb_UseClipboard.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chb_UseClipboard.Location = new System.Drawing.Point(16, 36);
            this.chb_UseClipboard.Name = "chb_UseClipboard";
            this.chb_UseClipboard.Size = new System.Drawing.Size(142, 17);
            this.chb_UseClipboard.TabIndex = 2;
            this.chb_UseClipboard.Text = "Copy results to Clipboard";
            this.chb_UseClipboard.UseVisualStyleBackColor = true;
            this.chb_UseClipboard.CheckedChanged += new System.EventHandler(this.chb_UseClipboard_CheckedChanged);
            // 
            // txt_TemplatePath
            // 
            this.txt_TemplatePath.Location = new System.Drawing.Point(95, 10);
            this.txt_TemplatePath.Name = "txt_TemplatePath";
            this.txt_TemplatePath.ReadOnly = true;
            this.txt_TemplatePath.Size = new System.Drawing.Size(222, 20);
            this.txt_TemplatePath.TabIndex = 1;
            this.txt_TemplatePath.DoubleClick += new System.EventHandler(this.txt_TemplatePath_DoubleClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Template Used";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 236);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(181, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Written By: Konstantinos Theodoridis";
            // 
            // lbl_github
            // 
            this.lbl_github.AutoSize = true;
            this.lbl_github.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.lbl_github.Location = new System.Drawing.Point(13, 252);
            this.lbl_github.Name = "lbl_github";
            this.lbl_github.Size = new System.Drawing.Size(158, 13);
            this.lbl_github.TabIndex = 5;
            this.lbl_github.Text = "https://github.com/ktheod";
            this.lbl_github.Click += new System.EventHandler(this.lbl_github_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(254, 252);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Version: 1.1";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(329, 275);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lbl_github);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.chb_UseClipboard);
            this.Controls.Add(this.txt_TemplatePath);
            this.Controls.Add(this.b_PasteClipBoard);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "NAV Object Version Fixer";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.mainForm_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button b_PasteClipBoard;
        private System.Windows.Forms.CheckBox chb_UseClipboard;
        private System.Windows.Forms.TextBox txt_TemplatePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lbl_github;
        private System.Windows.Forms.Label label4;
    }
}

