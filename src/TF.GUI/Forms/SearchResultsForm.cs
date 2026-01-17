using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using TF.Core.Entities;
using WeifenLuo.WinFormsUI.Docking;

namespace TF.GUI.Forms
{
	public partial class SearchResultsForm : DockContent
	{
		private class TupleView
		{
			public TranslationFileContainer Container;
			public TranslationFile File;

			public TupleView(Tuple<TranslationFileContainer, TranslationFile> tuple)
			{
				Container = tuple.Item1;
				File = tuple.Item2;
			}

			public override string ToString()
			{
				return $"{Path.Combine(Container.Path, File.RelativePath)}";
			}
		}

		public delegate bool FileChangedHandler(TranslationFile selectedFile);

		public event FileChangedHandler FileChanged;

		public SearchResultsForm()
		{
			InitializeComponent();
			AutoScaleMode = AutoScaleMode.Dpi;
		}

		protected virtual bool OnFileChanged(TranslationFile selectedFile)
		{
			if (FileChanged != null)
			{
				var cancel = FileChanged.Invoke(selectedFile);
				return cancel;
			}

			return false;
		}

		public void LoadItems(string searchString, IList<Tuple<TranslationFileContainer, TranslationFile>> filesFound)
		{
			lbSearchResult.Items.Clear();

			if (filesFound != null && filesFound.Count > 0)
			{
				string foundMessage = string.Format(LanguageManager.GetTranslation(2), searchString, filesFound.Count);
				lbSearchResult.Items.Add(foundMessage);

				foreach (var tuple in filesFound)
				{
					lbSearchResult.Items.Add(new TupleView(tuple));
				}
			}
			else
			{
				string notFoundMessage = string.Format(LanguageManager.GetTranslation(3), searchString);
				lbSearchResult.Items.Add(notFoundMessage);
			}
		}

		private void lbSearchResult_SelectedIndexChanged(object sender, EventArgs e)
		{
			var selectedItem = lbSearchResult.SelectedItem;
			if (selectedItem is TupleView tuple)
			{
				OnFileChanged(tuple.File);
			}
		}

		public void UpdateGUIText()
		{
			// Define translations: key = control/menu item reference, value = list of translations
			Dictionary<object, List<string>> translations = new Dictionary<object, List<string>>
			{
				{ this, new List<string> { "Risultati della ricerca", "Search results", "Suchergebnisse" } },
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
