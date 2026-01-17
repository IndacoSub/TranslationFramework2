using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection.Emit;
using System.Windows.Forms;
using WeifenLuo.WinFormsUI.Docking;

namespace TF.GUI.Forms
{
	public partial class WorkingForm : Form
	{
		private readonly BackgroundWorker _worker;
		private readonly bool _autoClose;

		public event DoWorkEventHandler DoWork;

		public bool Cancelled;
		protected WorkingForm()
		{
			InitializeComponent();
			AutoScaleMode = AutoScaleMode.Dpi;

			UpdateGUIText();

			_worker = new BackgroundWorker { WorkerReportsProgress = true, WorkerSupportsCancellation = true };

			_worker.RunWorkerCompleted += (sender, args) =>
			{
				Cancelled = args.Cancelled;

				if (_autoClose)
				{
					Close();
				}
				else
				{
					progressBar1.Visible = false;
					btnCancel.Visible = false;
					btnClose.Visible = true;
				}
			};

			_worker.ProgressChanged += (sender, args) => { AddProgress(args.UserState.ToString()); };
		}

		public WorkingForm(ThemeBase theme, string label, bool autoClose = false) : this()
		{
			dockPanel1.Theme = theme;
			Text = label;
			_autoClose = autoClose;
		}

		public void AddProgress(string text)
		{
			textBox1.AppendText(text + Environment.NewLine);
		}

		private void btnCancel_Click(object sender, EventArgs e)
		{
			string cancelMessage = LanguageManager.GetTranslation(4);
			_worker.ReportProgress(-1, cancelMessage);
			_worker.CancelAsync();

			btnCancel.Enabled = false;
		}

		private void WorkingForm_Shown(object sender, EventArgs e)
		{
			_worker.DoWork += this.DoWork;
			_worker.RunWorkerAsync();
		}

		private void btnClose_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		public void UpdateGUIText()
		{
			Dictionary<object, List<string>> translations = new Dictionary<object, List<string>>
			{
				{ btnClose, new List<string> { "Chiudi", "Close", "Schließen" } },
				{ btnCancel, new List<string> {"Annulla", "Cancel", "Abbrechen" } },
			};

			foreach (var pair in translations)
			{
				if (LanguageManager.LanguageIndex < pair.Value.Count)
				{
					switch (pair.Key)
					{
						case Control ctrl:
							ctrl.Text = pair.Value[LanguageManager.LanguageIndex];
							break;
						case ToolStripItem item:
							item.Text = pair.Value[LanguageManager.LanguageIndex];
							break;
						default:
							throw new InvalidOperationException("Unsupported UI element type.");
					}
				}
			}
		}
	}
}
