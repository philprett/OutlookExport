using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookExport
{
	public partial class Form1 : Form
	{
		Form1Controller controller;

		public Form1()
		{
			InitializeComponent();
			controller = new Form1Controller();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			txtExportFolder.Text = controller.LoadFormDetailsExportFolder();
			chkExportMsg.Checked = controller.LoadFormDetailsExportMsg();
			chkExportHtml.Checked = controller.LoadFormDetailsExportHtml();

			refreshFolderListToolStripMenuItem_Click(sender, e);
		}

		private void Form1_FormClosing(object sender, FormClosingEventArgs e)
		{
			List<string> selectedFolders = new List<string>();
			foreach (string f in lstFolders.SelectedItems) selectedFolders.Add(f);
			controller.SaveFormDetails(selectedFolders.ToArray(), txtExportFolder.Text, chkExportMsg.Checked, chkExportHtml.Checked);
		}

		private void refreshFolderListToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Cursor = Cursors.WaitCursor;

			controller.RefreshFolderList(lstFolders);

			string[] selectedFolders = controller.LoadFormDetailsSelectedFolders();
			foreach (string f in selectedFolders)
			{
				for (int a = 0; a < lstFolders.Items.Count; a++)
				{
					if ((string)lstFolders.Items[a] == f)
					{
						lstFolders.SetSelected(a, true);
					}
				}
			}

			Cursor = Cursors.Default;
		}

		private void butExportFolderBrowse_Click(object sender, EventArgs e)
		{
			controller.BrowseForExportFolder(txtExportFolder);
		}

		private void butExport_Click(object sender, EventArgs e)
		{
			lstFolders.Enabled = false;
			txtExportFolder.Enabled = false;
			butExportFolderBrowse.Enabled = false;
			chkExportMsg.Enabled = false;
			chkExportHtml.Enabled = false;
			butExport.Enabled = false;
			butCancel.Enabled = true;

			try
			{
				controller.Export(lstFolders, txtExportFolder, txtStatus, chkExportMsg.Checked, chkExportHtml.Checked, butCancel);
			}
			catch (Exception ex)
			{
				txtStatus.AppendText("ERROR\r\n");
				txtStatus.AppendText(ex.Message);
				txtStatus.AppendText("\r\n");
			}

			lstFolders.Enabled = true;
			txtExportFolder.Enabled = true;
			butExportFolderBrowse.Enabled = true;
			chkExportMsg.Enabled = true;
			chkExportHtml.Enabled = true;
			butExport.Enabled = true;
			butCancel.Enabled = false;
		}

		private void butCancel_Click(object sender, EventArgs e)
		{
			butCancel.Enabled = false;
		}
	}
}
