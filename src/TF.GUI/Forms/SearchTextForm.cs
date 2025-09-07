using System;
using System.Collections.Generic;
using System.Reflection.Emit;
using System.Windows.Forms;
using WeifenLuo.WinFormsUI.Docking;

namespace TF.GUI.Forms
{
	public partial class SearchTextForm : Form
	{
		public string SearchString => txtSearchString.Text;

		protected SearchTextForm()
		{
			InitializeComponent();
			AutoScaleMode = AutoScaleMode.Dpi;
		}

		public SearchTextForm(ThemeBase theme) : this()
		{
			dockPanel1.Theme = theme;
		}

		private void UpdateAcceptButton()
		{
			btnOK.Enabled = !string.IsNullOrEmpty(txtSearchString.Text);
		}

		private void txtSearchString_TextChanged(object sender, EventArgs e)
		{
			UpdateAcceptButton();
		}

		public void UpdateGUIText()
		{
			// Define translations: key = control/menu item reference, value = list of translations
			Dictionary<object, List<string>> translations = new Dictionary<object, List<string>>
			{
				{ this, new List<string> { "Cerca testo", "Search text", "Text suchen" } },
				{ label1, new List<string> { "Testo da cercare: ", "Text to search for: ", "Zu suchender Text: " } },
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
