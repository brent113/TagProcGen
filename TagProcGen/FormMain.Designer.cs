namespace TagProcGen
{
    partial class FormMain
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
            this._Gen = new System.Windows.Forms.Button();
            this._Browse = new System.Windows.Forms.Button();
            this._Label2 = new System.Windows.Forms.Label();
            this._Path = new System.Windows.Forms.TextBox();
            this._Label1 = new System.Windows.Forms.Label();
            this._OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this._Label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // _Gen
            // 
            this._Gen.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this._Gen.Location = new System.Drawing.Point(185, 103);
            this._Gen.Name = "_Gen";
            this._Gen.Size = new System.Drawing.Size(75, 23);
            this._Gen.TabIndex = 10;
            this._Gen.Text = "Generate";
            this._Gen.UseVisualStyleBackColor = true;
            this._Gen.Click += new System.EventHandler(this._Gen_Click);
            // 
            // _Browse
            // 
            this._Browse.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this._Browse.Location = new System.Drawing.Point(370, 70);
            this._Browse.Name = "_Browse";
            this._Browse.Size = new System.Drawing.Size(31, 23);
            this._Browse.TabIndex = 9;
            this._Browse.Text = "...";
            this._Browse.UseVisualStyleBackColor = true;
            this._Browse.Click += new System.EventHandler(this._Browse_Click);
            // 
            // _Label2
            // 
            this._Label2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this._Label2.AutoSize = true;
            this._Label2.Location = new System.Drawing.Point(42, 75);
            this._Label2.Name = "_Label2";
            this._Label2.Size = new System.Drawing.Size(26, 13);
            this._Label2.TabIndex = 8;
            this._Label2.Text = "File:";
            // 
            // _Path
            // 
            this._Path.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this._Path.Location = new System.Drawing.Point(74, 72);
            this._Path.Name = "_Path";
            this._Path.Size = new System.Drawing.Size(297, 20);
            this._Path.TabIndex = 7;
            // 
            // _Label1
            // 
            this._Label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._Label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this._Label1.Location = new System.Drawing.Point(12, 11);
            this._Label1.Name = "_Label1";
            this._Label1.Size = new System.Drawing.Size(396, 23);
            this._Label1.TabIndex = 6;
            this._Label1.Text = "Tag Processor Generator";
            this._Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // _Label3
            // 
            this._Label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._Label3.Location = new System.Drawing.Point(15, 32);
            this._Label3.Name = "_Label3";
            this._Label3.Size = new System.Drawing.Size(393, 35);
            this._Label3.TabIndex = 11;
            this._Label3.Text = "Takes the Excel file that contains the template definitions and creates the tag p" +
    "rocessor map.";
            this._Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // FormMain
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(420, 136);
            this.Controls.Add(this._Gen);
            this.Controls.Add(this._Browse);
            this.Controls.Add(this._Label2);
            this.Controls.Add(this._Path);
            this.Controls.Add(this._Label1);
            this.Controls.Add(this._Label3);
            this.MinimumSize = new System.Drawing.Size(436, 175);
            this.Name = "FormMain";
            this.Text = "TagProcGen";
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button _Gen;
        private System.Windows.Forms.Button _Browse;
        private System.Windows.Forms.Label _Label2;
        private System.Windows.Forms.TextBox _Path;
        private System.Windows.Forms.Label _Label1;
        private System.Windows.Forms.OpenFileDialog _OpenFileDialog1;
        private System.Windows.Forms.Label _Label3;
    }
}