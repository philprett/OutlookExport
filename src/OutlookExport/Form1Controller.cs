using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookExport
{
	class Form1Controller
	{
		private const int olMsg = 3;
		private const int olHtml = 5;

		private const int maxSubjectLength = 50;

		Microsoft.Office.Interop.Outlook.Application outlook;

		public Form1Controller()
		{
			outlook = new Microsoft.Office.Interop.Outlook.Application();
		}

		public void SaveFormDetails(string[] selectedFolders, string exportFolder, bool exportMsg, bool exportHtml)
		{
			string filename = string.Format("{0}.settings", Application.ExecutablePath);

			StringBuilder sb = new StringBuilder();
			sb.AppendLine(string.Format("ExportMsg={0}", exportMsg ? 1 : 0));
			sb.AppendLine(string.Format("ExportHtml={0}", exportHtml ? 1 : 0));
			sb.AppendLine(string.Format("ExportFolder={0}", exportFolder));
			foreach (string f in selectedFolders)
			{
				sb.AppendLine(string.Format("SelectedFolder={0}", f));
			}

			File.WriteAllText(filename, sb.ToString());
		}

		public string[] LoadFormDetailsSelectedFolders()
		{
			string filename = string.Format("{0}.settings", Application.ExecutablePath);

			List<string> selectedFolders = new List<string>();

			string details = File.Exists(filename) ? File.ReadAllText(filename) : string.Empty;
			foreach (string d in details.Split(new[] { '\n' }))
			{
				if (d.StartsWith("SelectedFolder="))
				{
					selectedFolders.Add(d.Substring(d.IndexOf("=") + 1).Trim());
				}
			}

			return selectedFolders.ToArray();
		}

		public string LoadFormDetailsExportFolder()
		{
			string filename = string.Format("{0}.settings", Application.ExecutablePath);

			string details = File.Exists(filename) ? File.ReadAllText(filename) : string.Empty;
			foreach (string d in details.Split(new[] { '\n' }))
			{
				if (d.StartsWith("ExportFolder="))
				{
					return d.Substring(d.IndexOf("=") + 1).Trim();
				}
			}

			return string.Empty;
		}

		public bool LoadFormDetailsExportMsg()
		{
			string filename = string.Format("{0}.settings", Application.ExecutablePath);

			string details = File.Exists(filename) ? File.ReadAllText(filename) : string.Empty;
			foreach (string d in details.Split(new[] { '\n' }))
			{
				if (d.StartsWith("ExportMsg="))
				{
					return d.Substring(d.IndexOf("=") + 1).Trim() == "1";
				}
			}

			return true;
		}

		public bool LoadFormDetailsExportHtml()
		{
			string filename = string.Format("{0}.settings", Application.ExecutablePath);

			string details = File.Exists(filename) ? File.ReadAllText(filename) : string.Empty;
			foreach (string d in details.Split(new[] { '\n' }))
			{
				if (d.StartsWith("ExportHtml="))
				{
					return d.Substring(d.IndexOf("=") + 1).Trim() == "1";
				}
			}

			return true;
		}

		public void RefreshFolderList(ListBox listBox, Microsoft.Office.Interop.Outlook.Folder folder = null)
		{
			if (folder == null)
			{
				List<Microsoft.Office.Interop.Outlook.Folder> folders = new List<Microsoft.Office.Interop.Outlook.Folder>();
				foreach (Microsoft.Office.Interop.Outlook.Folder f in outlook.Session.Folders)
				{
					folders.Add(f);
				}

				listBox.Items.Clear();
				Application.DoEvents();

				foreach (Microsoft.Office.Interop.Outlook.Folder f in folders.OrderBy(f => f.FolderPath))
				{
					listBox.Items.Add(f.FolderPath);
					Application.DoEvents();

					RefreshFolderList(listBox, f);
				}
			}
			else
			{
				List<Microsoft.Office.Interop.Outlook.Folder> folders = new List<Microsoft.Office.Interop.Outlook.Folder>();
				foreach (Microsoft.Office.Interop.Outlook.Folder f in folder.Folders)
				{
					folders.Add(f);
				}
				foreach (Microsoft.Office.Interop.Outlook.Folder f in folders.OrderBy(f => f.FolderPath))
				{
					listBox.Items.Add(f.FolderPath);
					Application.DoEvents();

					RefreshFolderList(listBox, f);
				}
			}
		}

		public void BrowseForExportFolder(TextBox textBox)
		{
			FolderBrowserDialog f = new FolderBrowserDialog();
			f.SelectedPath = textBox.Text;
			if (f.ShowDialog() == DialogResult.OK)
			{
				textBox.Text = f.SelectedPath;
				Application.DoEvents();
			}
		}

		public void Export(ListBox folders, TextBox exportFolder, TextBox status, bool exportMsg, bool exportHtml, Button cancelButton)
		{
			status.Text = "";
			Application.DoEvents();

			List<string> selectedFolders = new List<string>();
			foreach (string f in folders.SelectedItems) selectedFolders.Add(f);

			string exportFolderText = exportFolder.Text;

			foreach (Microsoft.Office.Interop.Outlook.Folder f in outlook.Session.Folders)
			{
				ExportFolder(
					selectedFolders.ToArray(),
					f,
					exportMsg ? exportFolderText + "\\msg" : string.Empty,
					exportHtml ? exportFolderText + "\\html" : string.Empty,
					status,
					cancelButton);
				if (!cancelButton.Enabled) return;
			}
		}

		private void ExportFolder(
			string[] selectedFolders,
			Microsoft.Office.Interop.Outlook.Folder outlookFolder,
			string exportFolderMsg,
			string exportFolderHtml,
			TextBox status,
			Button cancelButton)
		{
			string outlookFolderPath = outlookFolder.FolderPath;

			status.AppendText(string.Format("Scanning folder {0}...\r\n", outlookFolderPath));
			Application.DoEvents();

			// If the folder is selected, get the list of mails and export them
			if (selectedFolders.FirstOrDefault(f => outlookFolderPath.StartsWith(f)) != null)
			{
				if (!string.IsNullOrWhiteSpace(exportFolderMsg) && !Directory.Exists(exportFolderMsg))
				{
					Directory.CreateDirectory(exportFolderMsg);
				}
				if (!string.IsNullOrWhiteSpace(exportFolderHtml) && !Directory.Exists(exportFolderHtml))
				{
					Directory.CreateDirectory(exportFolderHtml);
				}

				string lastExportedMail = "";

				try
				{
					foreach (object item in outlookFolder.Items)
					{
						try
						{
							if (item is Microsoft.Office.Interop.Outlook.MailItem)
							{
								Microsoft.Office.Interop.Outlook.MailItem mail = (Microsoft.Office.Interop.Outlook.MailItem)item;


								string sentOn = mail.SentOn.ToString("yyyyMMdd_HHmmss");
								string sender = mail.SenderName;
								string subject = mail.Subject == null ? "" : mail.Subject.Length > maxSubjectLength ? mail.Subject.Substring(0, maxSubjectLength) : mail.Subject;

								string fullName = string.Format("{0} - {1} - {2}", sentOn, sender, subject);

								fullName = fullName.Replace("\\", "_");
								fullName = fullName.Replace("/", "_");
								fullName = fullName.Replace(":", "_");
								fullName = fullName.Replace("*", "_");
								fullName = fullName.Replace("?", "_");
								fullName = fullName.Replace("<", "_");
								fullName = fullName.Replace(">", "_");
								fullName = fullName.Replace("|", "_");
								fullName = fullName.Replace("\"", "_");

								if (!string.IsNullOrWhiteSpace(exportFolderMsg))
								{
									string exportFilenameMsg = string.Format("{0}\\{1}.msg", exportFolderMsg, fullName);
									if (!File.Exists(exportFilenameMsg))
									{
										status.AppendText(string.Format("    Saving Msg {0}...\r\n", fullName));
										Application.DoEvents();

										try
										{
											mail.SaveAs(exportFilenameMsg, olMsg);
										}
										catch (Exception ex)
										{
											status.AppendText(string.Format("ERROR Saving Msg {0}\r\n", fullName));
											status.AppendText(string.Format("    {0}\r\n", ex.Message));
											Application.DoEvents();
										}
										Application.DoEvents();
									}
								}

								if (!string.IsNullOrWhiteSpace(exportFolderHtml))
								{
									string exportFilenameHtml = string.Format("{0}\\{1}.txt", exportFolderHtml, fullName);
									if (!File.Exists(exportFilenameHtml))
									{
										status.AppendText(string.Format("    Saving Txt {0}...\r\n", fullName));
										Application.DoEvents();

										if (mail.Attachments.Count > 0)
										{
											string attachmentsFolder = string.Format("{0}\\.attachments\\{1}", exportFolderHtml, sentOn);
											if (!Directory.Exists(attachmentsFolder)) Directory.CreateDirectory(attachmentsFolder);

											for (int a = 1; a <= mail.Attachments.Count; a++)
											{
												string attachmentFilename;
												try
												{
													attachmentFilename = mail.Attachments[a].FileName;
												}
												catch
												{
													try
													{
														attachmentFilename = string.Format("{0} - {1}", a, mail.Attachments[a].DisplayName);
													}
													catch
													{
														attachmentFilename = string.Format("{0} - unknown type", a);
													}
												}
												try
												{
													attachmentFilename = string.Format("{0}\\{1}", attachmentsFolder, attachmentFilename);
													if (!File.Exists(attachmentFilename))
													{
														mail.Attachments[a].SaveAsFile(attachmentFilename);
													}
												}
												catch (Exception ex)
												{
													status.AppendText("    Error saving attachment ");
													status.AppendText(a.ToString());
													status.AppendText(" ");
													status.AppendText(attachmentFilename);
													status.AppendText("\r\n");
													throw ex;
												}
											}
										}

										mail.SaveAs(exportFilenameHtml, Microsoft.Office.Interop.Outlook.OlSaveAsType.olTXT);
										Application.DoEvents();
									}
									else
									{
										//status.AppendText("    Already exists: ");
										//status.AppendText(exportFilenameHtml);
										//status.AppendText("\r\n");
									}
								}
								mail = null;

								lastExportedMail = fullName;
							}
						}
						catch (Exception ex)
						{
							status.AppendText("ERROR exporting mail in ");
							status.AppendText(outlookFolder.FolderPath);
							status.AppendText("\r\n");
							status.AppendText(ex.Message);
							status.AppendText("\r\n");

							status.AppendText("Last exported mail: ");
							status.AppendText(lastExportedMail);
							status.AppendText("\r\n");
						}
						if (!cancelButton.Enabled) return;

						Application.DoEvents();
					}
				}
				catch (Exception ex)
				{
					status.AppendText("ERROR exporting folder ");
					status.AppendText(outlookFolder.FolderPath);
					status.AppendText("\r\n");

					status.AppendText("Last exported mail: ");
					status.AppendText(lastExportedMail);
					status.AppendText("\r\n");

					throw ex;
				}
				GC.Collect();
			}


			// Now process and subfolders
			foreach (Microsoft.Office.Interop.Outlook.Folder f in outlookFolder.Folders)
			{
				ExportFolder(
					selectedFolders,
					f,
					!string.IsNullOrWhiteSpace(exportFolderMsg) ? string.Format("{0}\\{1}", exportFolderMsg, f.Name) : "",
					!string.IsNullOrWhiteSpace(exportFolderHtml) ? string.Format("{0}\\{1}", exportFolderHtml, f.Name) : "",
					status,
					cancelButton);
			}
		}
	}
}
