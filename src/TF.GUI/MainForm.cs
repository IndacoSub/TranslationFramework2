using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using TF.Core;

namespace TF.GUI
{

	public partial class MainForm : Form
	{
		private PluginManager _pluginManager;

		public MainForm()
		{
			InitializeComponent();
			LanguageManager.ApplyCulture();

			AutoScaleMode = AutoScaleMode.Dpi;
			this.LanguageComboBox.SelectedIndex = 0;
			this.LanguageComboBox.SelectedIndexChanged += new System.EventHandler(this.toolStripComboBox1_SelectedIndexChanged);

			ConfigureDock();

			_pluginManager = new PluginManager();
			_pluginManager.LoadPlugins(Path.Combine(Application.StartupPath, "plugins"));

			Core.Fonts.FontCollection.Initialize();
		}

		private void MainForm_Shown(object sender, EventArgs e)
		{
			if (_pluginManager.GetAllGames().Count == 0)
			{
				MessageBox.Show("No se ha podido encontrar ningún plugin válido.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Close();
			}
		}

		private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (!CloseAllDocuments())
			{
				e.Cancel = true;
				return;
			}
			SaveSettings();
		}

		private void mniFileExit_Click(object sender, EventArgs e)
		{
			Close();
		}

		private void FileNew_Click(object sender, EventArgs e)
		{
			CreateNewTranslation();
		}

		private void FileOpen_Click(object sender, EventArgs e)
		{
			LoadTranslation();
		}

		private void FileSave_Click(object sender, EventArgs e)
		{
			SaveChanges();
		}

		private void HelpAbout_Click(object sender, EventArgs e)
		{

		}

		private void FileExport_Click(object sender, EventArgs e)
		{
			ExportProject();
		}

		private void SearchInFiles_Click(object sender, EventArgs e)
		{
			SearchInFiles();
		}

		private void SearchText_Click(object sender, EventArgs e)
		{
			SearchText();
		}

		private void MainForm_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				if (e.Shift)
				{
					SearchText(-1);
				}
				else
				{
					SearchText(1);
				}
			}

			if (e.KeyCode == Keys.F8)
			{
				_explorer.SelectPrevious();
			}

			if (e.KeyCode == Keys.F9)
			{
				_explorer.SelectNext();
			}
		}

		private void MainForm_KeyUp(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				if (e.Shift)
				{
					SearchText(-1);
				}
				else
				{
					SearchText(1);
				}
			}

			if (e.KeyCode == Keys.F8)
			{
				_explorer.SelectPrevious();
			}

			if (e.KeyCode == Keys.F9)
			{
				_explorer.SelectNext();
			}
		}

		private void mniBulkTextsExportPo_Click(object sender, EventArgs e)
		{
			ExportTexts();
		}

		private void mniBulkTextsImportPo_Click(object sender, EventArgs e)
		{
			ImportTexts();
		}

		private void mniBulkImagesExport_Click(object sender, EventArgs e)
		{
			ExportImages();
		}

		private void mniBulkImagesImport_Click(object sender, EventArgs e)
		{
			ImportImages();
		}

		private void mniBulkTextsExportXlsx_Click(object sender, EventArgs e)
		{
			ExportTextsXlsx();
		}

		private void mniBulkTextsImportXlsx_Click(object sender, EventArgs e)
		{
			ImportTextsXlsxSimple();
		}

		private void mniBulkTextsImportXlsxOffset_Click(object sender, EventArgs e)
		{
			ImportTextsXlsxOffset();
		}

		private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{

			UpdateGUIText();
		}

		public void UpdateGUIText()
		{
			int selectedLanguageIndex = this.LanguageComboBox.SelectedIndex;
			LanguageManager.LanguageIndex = selectedLanguageIndex;

			_explorer.UpdateGUIText();
			_searchResults.UpdateGUIText();

			// Define translations: key = control/menu item reference, value = list of translations
			Dictionary<object, List<string>> translations = new Dictionary<object, List<string>>
			{
				{ this.mniFile, new List<string> { "File (A)", "Files (A)", "Datei (A)"} },
				{ this.mniEdit, new List<string> { "Modifica (E)", "Edit (E)", "Bearbeiten (E)" } },
				{ this.mniEditSearch, new List<string> { "Cerca (B)", "Search (B)", "Suchen (B)" } },
				{ this.mniEditSearchInFiles, new List<string> { "Cerca nei file (&u)", "Search in files (&u)", "In Dateien suchen (&u)" } },

				// File menu items
				{ this.mniFileNew, new List<string> { "&Nuovo", "&New", "&Neu" } },
				{ this.mniFileOpen, new List<string> { "&Apri", "&Open", "&Öffnen" } },
				{ this.mniFileSave, new List<string> { "Salva (&G)", "Save (&G)", "Speichern (&G)" } },
				{ this.mniFileExport, new List<string> { "Esporta...", "Export...", "Exportieren..." } },
				{ this.mniFileExit, new List<string> { "Esci (&S)", "Exit (&S)", "Beenden (&S)" } },

				{ this.mniBulk, new List<string> { "Massa", "Bulk", "Masse" } },
				{ this.mniBulkTexts, new List<string> { "Testi", "Texts", "Texte" } },
				{ this.mniBulkTextsExportPo, new List<string> { "Esporta (*.po)", "Export (*.po)", "Exportieren (*.po)" } },
				{ this.mniBulkTextsImportPo, new List<string> { "Importa (*.po)", "Import (*.po)", "Importieren (*.po)" } },
				{ this.mniBulkTextsExportXlsx, new List<string> { "Esporta (*.xlsx)", "Export (*.xlsx)", "Exportieren (*.xlsx)" } },
				{ this.mniBulkTextsImportXlsx, new List<string> { "Importa (*.xlsx)", "Import (*.xlsx)", "Importieren (*.xlsx)" } },
				{ this.mniBulkTextsImportXlsxOffset, new List<string> { "Importa (*.xlsx, Offset)", "Import (*.xlsx, Offset)", "Importieren (*.xlsx, Offset)" } },

				{ this.mniBulkImages, new List<string> { "Immagini", "Images", "Bilder" } },
				{ this.mniBulkImagesExport, new List<string> { "Esporta", "Export", "Exportieren" } },
				{ this.mniBulkImagesImport, new List<string> { "Importa", "Import", "Importieren" } },

				{ this.mniHelp, new List<string> { "Aiuto", "Help", "Hilfe" } },
				{ this.mniHelpAbout, new List<string> { "Cerc&a", "About", "Über" } },

				{ this.tsbSearchInFiles, new List<string> { "&Cerca nei file", "&Search in files", "&In Dateien suchen" } },
				{ this.tsbSearch, new List<string> { "&Cerca", "&Search", "&Suchen" } },
				{ this.tsbExportProject, new List<string> { "&Esporta", "&Export", "&Exportieren" } },
				{ this.tsbSaveFile, new List<string> { "&Salva", "&Save", "&Speichern" } },
				{ this.tsbOpenFile, new List<string> { "&Apri", "&Open", "&Öffnen" } },
				{ this.tsbNewFile, new List<string> { "&Nuovo", "&New", "&Neu" } },
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

	public static class LanguageManager
	{
		public static int LanguageIndex { get; set; } = 0;

		public static void ApplyCulture()
		{
			string[] cultures = { "it", "en", "de" };
			var culture = new CultureInfo(cultures[LanguageIndex]);
			Thread.CurrentThread.CurrentUICulture = culture;
			Thread.CurrentThread.CurrentCulture = culture;
		}

		public static string GetTranslation(uint index)
		{
			if (translations.TryGetValue(index, out var list) && LanguageIndex < list.Count)
			{
				return list[LanguageIndex];
			}

			return "[Translation missing]";
		}

		public static readonly Dictionary<uint, List<string>> translations = new Dictionary<uint, List<string>>
		{
			{ 0, new List<string> { string.Concat("Seleziona la cartella di lavoro.", Environment.NewLine, "I file necessari per la traduzione verranno salvati in essa.", Environment.NewLine, "Assicurati che ci sia spazio sufficiente sul disco."), string.Concat("Select the working folder.", Environment.NewLine, "The files needed for translation will be saved there.", Environment.NewLine, "Make sure there is enough disk space."), string.Concat("Wähle den Arbeitsordner aus.", Environment.NewLine, "Die für die Übersetzung benötigten Dateien werden dort gespeichert.", Environment.NewLine, "Stelle sicher, dass genügend Speicherplatz vorhanden ist.") } },
			{ 1, new List<string> { "Seleziona la cartella di installazione del gioco.", "Select the game installation folder.", "Wähle den Installationsordner des Spiels aus." } },
			{ 2, new List<string> { "Il testo \"{0}\" appare in {1} file del progetto", "The text \"{0}\" appears in {1} project files", "Der Text \"{0}\" erscheint in {1} Projektdateien" } },
			{ 3, new List<string> { "Il testo \"{0}\" non appare in alcun file del progetto", "The text \"{0}\" does not appear in any project file", "Der Text \"{0}\" erscheint in keiner Projektdatei" } },
			{ 4, new List<string> { "Annullamento in corso (attendere il completamento dell'attività corrente)...", "Cancelling (please wait for the current task to finish)...", "Abbruch läuft (bitte warten, bis die aktuelle Aufgabe abgeschlossen ist)..." } },
			{ 5, new List<string> { "Nuova traduzione", "New translation", "Neue Übersetzung" } },
			{ 6, new List<string> { "FINE", "DONE", "FERTIG" } },
			{ 7, new List<string> { "Eliminazione dei file...", "Deleting files...", "Dateien werden gelöscht..." } },
			{ 8, new List<string> { "Completata", "Completed", "Abgeschlossen" } },
			{ 9, new List<string> { "La cartella {0} non e' vuota. Selezionane una vuota.", "The folder {0} is not empty. Please select an empty one.", "Der Ordner {0} ist nicht leer. Bitte wähle einen leeren Ordner aus." } },
			{ 10, new List<string> { "Attenzione", "Warning", "Achtung" } },
			{ 11, new List<string> { "Carica traduzione", "Load translation", "Übersetzung laden" } },
			{ 12, new List<string> { "E' necessario salvare i cambiamenti prima di continuare.\nDesideri salvarli?", "You need to save changes before continuing.\nDo you want to save them?", "Sie müssen die Änderungen speichern, bevor Sie fortfahren.\nMöchten Sie sie speichern?" } },
			{ 13, new List<string> { "Salva cambiamenti", "Save changes", "Änderungen speichern" } },
			{ 14, new List<string> { "Esporta traduzione", "Export translation", "Übersetzung exportieren" } },
			{ 15, new List<string> { "Puoi trovare i file esportati nella cartella: {0}", "You can find the exported files in the folder: {0}", "Du findest die exportierten Dateien im Ordner: {0}" } },
			{ 16, new List<string> { "Cerca nei file", "Search in files", "In Dateien suchen" } },
			{ 17, new List<string> { "Nessuna corrispondenza trovata.", "No matches found.", "Keine Übereinstimmungen gefunden." } },
			{ 18, new List<string> { "Cerca", "Search", "Suche" } },
			{ 19, new List<string> { "Seleziona la cartella dove desideri salvare i/il file Po", "Select the folder where you want to save the Po file(s)", "Wähle den Ordner, in dem du die Po-Datei(en) speichern möchtest" } },
			{ 20, new List<string> { "Esporta file Po", "Export Po files", "Po-Dateien exportieren" } },
			{ 21, new List<string> { "Puoi trovare i file esportati nella cartella: {0}", "You can find the exported files in the folder: {0}", "Du findest die exportierten Dateien im Ordner: {0}" } },
			{ 22, new List<string> { "Seleziona la cartella dove desideri salvare le immagini.", "Select the folder where you want to save the images.", "Wähle den Ordner, in dem du die Bilder speichern möchtest." } },
			{ 23, new List<string> { "Esporta immagini", "Export images", "Bilder exportieren" } },
			{ 24, new List<string> { "Seleziona la cartella root contenente i file Po", "Select the root folder containing the Po files", "Wähle den Stammordner mit den Po-Dateien aus" } },
			{ 25, new List<string> { "Importa Po", "Import Po", "Po importieren" } },
			{ 26, new List<string> { "Seleziona la cartella dove desideri salvare i/il file Xlsx", "Select the folder where you want to save the Xlsx file(s)", "Wähle den Ordner, in dem du die Xlsx-Datei(en) speichern möchtest" } },
			{ 27, new List<string> { "Esporta file Xlsx", "Export Xlsx file(s)", "Xlsx-Datei(en) exportieren" } },
			{ 28, new List<string> { "Seleziona la cartella root contenente i file XLSX", "Select the root folder containing the XLSX files", "Wähle den Stammordner mit den XLSX-Dateien" } },
			{ 29, new List<string> { "Importa XLSX", "Import XLSX", "XLSX importieren" } },
			{ 30, new List<string> { "Importa XLSX (Offset)", "Import XLSX (Offset)", "XLSX (Offset) importieren" } },
			{ 31, new List<string> { "Seleziona la cartella root con le immagini", "Select the root folder with the images", "Wähle den Stammordner mit den Bildern" } },
			{ 32, new List<string> { "Importa immagini", "Import images", "Bilder importieren" } }
		};
	}
}
