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
		public Form1()
		{
			InitializeComponent();
		}

		private void butStartExport_Click(object sender, EventArgs e)
		{
			Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();

			foreach (var folder in outlook.Session.Folders)
			{

			}
		}

		private void Form1_Load(object sender, EventArgs e)
		{

		}
	}
}
