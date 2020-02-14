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
			this.menuStrip1 = new System.Windows.Forms.MenuStrip();
			this.refreshFolderListToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.label1 = new System.Windows.Forms.Label();
			this.lstFolders = new System.Windows.Forms.ListBox();
			this.label2 = new System.Windows.Forms.Label();
			this.txtExportFolder = new System.Windows.Forms.TextBox();
			this.butExportFolderBrowse = new System.Windows.Forms.Button();
			this.butExport = new System.Windows.Forms.Button();
			this.txtStatus = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.chkExportMsg = new System.Windows.Forms.CheckBox();
			this.chkExportHtml = new System.Windows.Forms.CheckBox();
			this.butCancel = new System.Windows.Forms.Button();
			this.menuStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			this.SuspendLayout();
			// 
			// menuStrip1
			// 
			this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshFolderListToolStripMenuItem});
			this.menuStrip1.Location = new System.Drawing.Point(0, 0);
			this.menuStrip1.Name = "menuStrip1";
			this.menuStrip1.Size = new System.Drawing.Size(800, 24);
			this.menuStrip1.TabIndex = 1;
			this.menuStrip1.Text = "menuStrip1";
			// 
			// refreshFolderListToolStripMenuItem
			// 
			this.refreshFolderListToolStripMenuItem.Name = "refreshFolderListToolStripMenuItem";
			this.refreshFolderListToolStripMenuItem.Size = new System.Drawing.Size(115, 20);
			this.refreshFolderListToolStripMenuItem.Text = "Refresh Folder List";
			this.refreshFolderListToolStripMenuItem.Click += new System.EventHandler(this.refreshFolderListToolStripMenuItem_Click);
			// 
			// splitContainer1
			// 
			this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.splitContainer1.Location = new System.Drawing.Point(0, 24);
			this.splitContainer1.Name = "splitContainer1";
			// 
			// splitContainer1.Panel1
			// 
			this.splitContainer1.Panel1.Controls.Add(this.lstFolders);
			this.splitContainer1.Panel1.Controls.Add(this.label1);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.butCancel);
			this.splitContainer1.Panel2.Controls.Add(this.chkExportHtml);
			this.splitContainer1.Panel2.Controls.Add(this.chkExportMsg);
			this.splitContainer1.Panel2.Controls.Add(this.label3);
			this.splitContainer1.Panel2.Controls.Add(this.txtStatus);
			this.splitContainer1.Panel2.Controls.Add(this.butExport);
			this.splitContainer1.Panel2.Controls.Add(this.butExportFolderBrowse);
			this.splitContainer1.Panel2.Controls.Add(this.txtExportFolder);
			this.splitContainer1.Panel2.Controls.Add(this.label2);
			this.splitContainer1.Size = new System.Drawing.Size(800, 426);
			this.splitContainer1.SplitterDistance = 234;
			this.splitContainer1.TabIndex = 2;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(3, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 13);
			this.label1.TabIndex = 0;
			this.label1.Text = "Folders";
			// 
			// lstFolders
			// 
			this.lstFolders.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.lstFolders.FormattingEnabled = true;
			this.lstFolders.Location = new System.Drawing.Point(3, 25);
			this.lstFolders.Name = "lstFolders";
			this.lstFolders.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.lstFolders.Size = new System.Drawing.Size(228, 394);
			this.lstFolders.TabIndex = 1;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(3, 9);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(69, 13);
			this.label2.TabIndex = 0;
			this.label2.Text = "Export Folder";
			// 
			// txtExportFolder
			// 
			this.txtExportFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtExportFolder.Location = new System.Drawing.Point(6, 25);
			this.txtExportFolder.Name = "txtExportFolder";
			this.txtExportFolder.Size = new System.Drawing.Size(463, 20);
			this.txtExportFolder.TabIndex = 1;
			// 
			// butExportFolderBrowse
			// 
			this.butExportFolderBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.butExportFolderBrowse.Location = new System.Drawing.Point(475, 23);
			this.butExportFolderBrowse.Name = "butExportFolderBrowse";
			this.butExportFolderBrowse.Size = new System.Drawing.Size(75, 23);
			this.butExportFolderBrowse.TabIndex = 2;
			this.butExportFolderBrowse.Text = "Browse";
			this.butExportFolderBrowse.UseVisualStyleBackColor = true;
			this.butExportFolderBrowse.Click += new System.EventHandler(this.butExportFolderBrowse_Click);
			// 
			// butExport
			// 
			this.butExport.Location = new System.Drawing.Point(222, 51);
			this.butExport.Name = "butExport";
			this.butExport.Size = new System.Drawing.Size(75, 23);
			this.butExport.TabIndex = 3;
			this.butExport.Text = "Export";
			this.butExport.UseVisualStyleBackColor = true;
			this.butExport.Click += new System.EventHandler(this.butExport_Click);
			// 
			// txtStatus
			// 
			this.txtStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtStatus.Location = new System.Drawing.Point(6, 102);
			this.txtStatus.Multiline = true;
			this.txtStatus.Name = "txtStatus";
			this.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.txtStatus.Size = new System.Drawing.Size(544, 317);
			this.txtStatus.TabIndex = 4;
			this.txtStatus.WordWrap = false;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(3, 86);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 13);
			this.label3.TabIndex = 5;
			this.label3.Text = "Status";
			// 
			// chkExportMsg
			// 
			this.chkExportMsg.AutoSize = true;
			this.chkExportMsg.Location = new System.Drawing.Point(6, 55);
			this.chkExportMsg.Name = "chkExportMsg";
			this.chkExportMsg.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.chkExportMsg.Size = new System.Drawing.Size(70, 17);
			this.chkExportMsg.TabIndex = 6;
			this.chkExportMsg.Text = "Msg Files";
			this.chkExportMsg.UseVisualStyleBackColor = true;
			// 
			// chkExportHtml
			// 
			this.chkExportHtml.AutoSize = true;
			this.chkExportHtml.Location = new System.Drawing.Point(97, 55);
			this.chkExportHtml.Name = "chkExportHtml";
			this.chkExportHtml.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.chkExportHtml.Size = new System.Drawing.Size(71, 17);
			this.chkExportHtml.TabIndex = 7;
			this.chkExportHtml.Text = "Text Files";
			this.chkExportHtml.UseVisualStyleBackColor = true;
			// 
			// butCancel
			// 
			this.butCancel.Enabled = false;
			this.butCancel.Location = new System.Drawing.Point(331, 51);
			this.butCancel.Name = "butCancel";
			this.butCancel.Size = new System.Drawing.Size(75, 23);
			this.butCancel.TabIndex = 8;
			this.butCancel.Text = "Cancel";
			this.butCancel.UseVisualStyleBackColor = true;
			this.butCancel.Click += new System.EventHandler(this.butCancel_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(800, 450);
			this.Controls.Add(this.splitContainer1);
			this.Controls.Add(this.menuStrip1);
			this.MainMenuStrip = this.menuStrip1;
			this.Name = "Form1";
			this.Text = "Outlook Export";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
			this.Load += new System.EventHandler(this.Form1_Load);
			this.menuStrip1.ResumeLayout(false);
			this.menuStrip1.PerformLayout();
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel1.PerformLayout();
			this.splitContainer1.Panel2.ResumeLayout(false);
			this.splitContainer1.Panel2.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion
		private System.Windows.Forms.MenuStrip menuStrip1;
		private System.Windows.Forms.ToolStripMenuItem refreshFolderListToolStripMenuItem;
		private System.Windows.Forms.SplitContainer splitContainer1;
		private System.Windows.Forms.ListBox lstFolders;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtStatus;
		private System.Windows.Forms.Button butExport;
		private System.Windows.Forms.Button butExportFolderBrowse;
		private System.Windows.Forms.TextBox txtExportFolder;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.CheckBox chkExportHtml;
		private System.Windows.Forms.CheckBox chkExportMsg;
		private System.Windows.Forms.Button butCancel;
	}
}

