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
            this.Gen = new System.Windows.Forms.Button();
            this.Browse = new System.Windows.Forms.Button();
            this.Label2 = new System.Windows.Forms.Label();
            this.Path = new System.Windows.Forms.TextBox();
            this.Label1 = new System.Windows.Forms.Label();
            this._OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.Label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Gen
            // 
            this.Gen.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.Gen.Location = new System.Drawing.Point(185, 103);
            this.Gen.Name = "Gen";
            this.Gen.Size = new System.Drawing.Size(75, 23);
            this.Gen.TabIndex = 10;
            this.Gen.Text = "Generate";
            this.Gen.UseVisualStyleBackColor = true;
            this.Gen.Click += new System.EventHandler(this.Gen_Click);
            // 
            // Browse
            // 
            this.Browse.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.Browse.Location = new System.Drawing.Point(370, 70);
            this.Browse.Name = "Browse";
            this.Browse.Size = new System.Drawing.Size(31, 23);
            this.Browse.TabIndex = 9;
            this.Browse.Text = "...";
            this.Browse.UseVisualStyleBackColor = true;
            this.Browse.Click += new System.EventHandler(this.Browse_Click);
            // 
            // Label2
            // 
            this.Label2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(42, 75);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(26, 13);
            this.Label2.TabIndex = 8;
            this.Label2.Text = "File:";
            // 
            // Path
            // 
            this.Path.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Path.Location = new System.Drawing.Point(74, 72);
            this.Path.Name = "Path";
            this.Path.Size = new System.Drawing.Size(297, 20);
            this.Path.TabIndex = 7;
            // 
            // Label1
            // 
            this.Label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label1.Location = new System.Drawing.Point(12, 11);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(396, 23);
            this.Label1.TabIndex = 6;
            this.Label1.Text = "Tag Processor Generator";
            this.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label3
            // 
            this.Label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Label3.Location = new System.Drawing.Point(15, 32);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(393, 35);
            this.Label3.TabIndex = 11;
            this.Label3.Text = "Takes the Excel file that contains the template definitions and creates the tag p" +
    "rocessor map.";
            this.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // FormMain
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(420, 136);
            this.Controls.Add(this.Gen);
            this.Controls.Add(this.Browse);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.Path);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.Label3);
            this.MinimumSize = new System.Drawing.Size(436, 175);
            this.Name = "FormMain";
            this.Text = "TagProcGen";
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.FormMain_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.FormMain_DragEnter);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Gen;
        private System.Windows.Forms.Button Browse;
        private System.Windows.Forms.Label Label2;
        private System.Windows.Forms.TextBox Path;
        private System.Windows.Forms.Label Label1;
        private System.Windows.Forms.OpenFileDialog _OpenFileDialog1;
        private System.Windows.Forms.Label Label3;
    }
}