namespace OutlookExport
{
	partial class Form1
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
			this.butStartExport = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// butStartExport
			// 
			this.butStartExport.Location = new System.Drawing.Point(98, 125);
			this.butStartExport.Name = "butStartExport";
			this.butStartExport.Size = new System.Drawing.Size(92, 38);
			this.butStartExport.TabIndex = 0;
			this.butStartExport.Text = "Start Export";
			this.butStartExport.UseVisualStyleBackColor = true;
			this.butStartExport.Click += new System.EventHandler(this.butStartExport_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(800, 450);
			this.Controls.Add(this.butStartExport);
			this.Name = "Form1";
			this.Text = "Outlook Export";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button butStartExport;
	}
}

