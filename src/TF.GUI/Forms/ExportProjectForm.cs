using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Windows.Forms;
using TF.Core.Entities;
using TF.Core.POCO;
using WeifenLuo.WinFormsUI.Docking;

namespace TF.GUI.Forms
{
	public partial class ExportProjectForm : Form
	{
		public IList<TranslationFileContainer> SelectedContainers
		{
			get
			{
				var result = new List<TranslationFileContainer>();
				for (var i = 0; i < lbItems.Items.Count; i++)
				{
					var isChecked = lbItems.GetItemChecked(i);
					if (isChecked)
					{
						result.Add(lbItems.Items[i] as TranslationFileContainer);
					}
				}

				return result;
			}
		}

		public ExportOptions Options
		{
			get
			{
				var result = new ExportOptions
				{
					UseCompression = chkCompress.Checked,
					ForceRebuild = chkForceRebuild.Checked,
					SaveTempFiles = chkSaveTempFiles.Checked
				};
				return result;
			}
		}

		protected ExportProjectForm()
		{
			InitializeComponent();
			AutoScaleMode = AutoScaleMode.Dpi;
		}

		public ExportProjectForm(ThemeBase theme, IList<TranslationFileContainer> containers) : this()
		{
			dockPanel1.Theme = theme;

			lbItems.Items.Clear();
			foreach (var container in containers)
			{
				lbItems.Items.Add(container, container.Files.Any(x => x.HasChanges));
			}
		}

		private void btnSelectAll_Click(object sender, EventArgs e)
		{
			for (var i = 0; i < lbItems.Items.Count; i++)
			{
				lbItems.SetItemChecked(i, true);
			}
		}

		private void btnSelectNone_Click(object sender, EventArgs e)
		{
			for (var i = 0; i < lbItems.Items.Count; i++)
			{
				lbItems.SetItemChecked(i, false);
			}
		}

		private void btnInvertSelection_Click(object sender, EventArgs e)
		{
			for (var i = 0; i < lbItems.Items.Count; i++)
			{
				var currentState = lbItems.GetItemChecked(i);
				lbItems.SetItemChecked(i, !currentState);
			}
		}

		private void lbItems_ItemCheck(object sender, ItemCheckEventArgs e)
		{
			for (var i = 0; i < lbItems.Items.Count; i++)
			{
				bool currentState;
				if (i == e.Index)
				{
					currentState = e.NewValue == CheckState.Checked;
				}
				else
				{
					currentState = lbItems.GetItemChecked(i);
				}

				if (currentState)
				{
					btnOK.Enabled = true;
					return;
				}
			}

			btnOK.Enabled = false;
		}

		private void btnOnlyModified_Click(object sender, EventArgs e)
		{
			for (var i = 0; i < lbItems.Items.Count; i++)
			{
				var isChecked = ((TranslationFileContainer)lbItems.Items[i]).Files.Any(x => x.HasChanges);
				lbItems.SetItemChecked(i, isChecked);
			}
		}

		public void UpdateGUIText()
		{
			// Define translations: key = control/menu item reference, value = list of translations
			Dictionary<object, List<string>> translations = new Dictionary<object, List<string>>
			{
				{ this, new List<string> { "Nuova Traduzione", "New Translation", "Neue Übersetzung" } },
				{ chkCompress, new List<string> { "Comprimi", "Compress", "Komprimieren" } },
				{ chkForceRebuild, new List<string> { "Forza creazione file", "Force rebuild", "Erzwingen Sie einen Neuaufbau" } },
				{ gbExportItems, new List<string> { "Elementi da esportare", "Items to export", "Zu exportierende Elemente" } },
				{ btnSelectAll, new List<string> { "Tutti", "All", "Alle" } },
				{ btnSelectNone, new List<string> { "Nessuno", "None", "Kein" } },
				{ btnInvertSelection, new List<string> { "Inverti selezione", "Reverse selection", "Auswahl umkehren" } },
				{ btnOnlyModified, new List<string> { "Solo modificati", "Only modified", "Nur geänderte" } },
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
