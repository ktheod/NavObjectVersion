﻿namespace NAVObjectVersion
{
    partial class mainForm
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
            this.b_PasteClipBoard = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // b_PasteClipBoard
            // 
            this.b_PasteClipBoard.Location = new System.Drawing.Point(13, 13);
            this.b_PasteClipBoard.Name = "b_PasteClipBoard";
            this.b_PasteClipBoard.Size = new System.Drawing.Size(161, 138);
            this.b_PasteClipBoard.TabIndex = 0;
            this.b_PasteClipBoard.Text = "Paste Clipboard";
            this.b_PasteClipBoard.UseVisualStyleBackColor = true;
            this.b_PasteClipBoard.Click += new System.EventHandler(this.b_LoadClipboard_Click);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(187, 164);
            this.Controls.Add(this.b_PasteClipBoard);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "mainForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "NAV Object Version Fixer";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button b_PasteClipBoard;
    }
}

