﻿using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using TF.Core.Entities;
using TF.Core.Exceptions;
using TF.Core.Helpers;
using TF.Core.POCO;
using TF.Core.TranslationEntities;
using TF.IO;
using WeifenLuo.WinFormsUI.Docking;
using System.Text.RegularExpressions;
using TF.Core.Files;
using TF.Core.Views;

namespace TF.Core
{
    public class TranslationProject
    {
        public IGame Game { get; private set; }
        public string InstallationPath { get; private set; }
        public string WorkPath { get; private set; }

        public string ContainersFolder => Path.Combine(WorkPath, "containers");
        public string ChangesFolder => Path.Combine(WorkPath, "changes");
        public string ExportFolder => Path.Combine(WorkPath, "export");
        public string TempFolder => Path.Combine(WorkPath, "temp");

        public IList<TranslationFileContainer> FileContainers { get; private set; }

        private TranslationProject()
        {
            FileContainers = new List<TranslationFileContainer>();
        }

        public TranslationProject(IGame game, string installFolder, string path) : this()
        {
            Game = game;
            InstallationPath = installFolder;
            WorkPath = path;

            Directory.CreateDirectory(ContainersFolder);
            Directory.CreateDirectory(ChangesFolder);
            Directory.CreateDirectory(ExportFolder);
            Directory.CreateDirectory(TempFolder);
        }

        public void ReadTranslationFiles(BackgroundWorker worker)
        {
            UpdateTranslationFiles(this, worker);
        }

        public void Save()
        {
            string saveFile = Path.Combine(WorkPath, "project.tf");
            using (var fs = new FileStream(saveFile, FileMode.Create))
            using (var output = new ExtendedBinaryWriter(fs, Encoding.UTF8))
            {
                output.WriteString(Game.Id);
                output.Write(Game.Version);
                output.WriteString(InstallationPath);
                output.Write(FileContainers.Count);
                foreach (TranslationFileContainer container in FileContainers)
                {
                    output.WriteString(container.Id);
                    output.WriteString(container.Path);
                    output.Write((int)container.Type);
                    output.Write(container.Files.Count);
                    foreach (TranslationFile file in container.Files)
                    {
                        output.WriteString(file.Id);
                        output.WriteString(file.Path);
                        output.WriteString(file.RelativePath);
                        output.WriteString(file.Name);
                        output.WriteString(file.GetType().FullName);
                    }
                }
            }
        }

        public static TranslationProject Load(string path, PluginManager pluginManager, BackgroundWorker worker)
        {
            var types = new Dictionary<string, Type>();

            var searchNewFiles = false;

            using (var fs = new FileStream(path, FileMode.Open))
            using (var input = new ExtendedBinaryReader(fs, Encoding.UTF8))
            {
                var result = new TranslationProject();

                string gameId = input.ReadString();
                IGame game = pluginManager.GetGame(gameId);
                result.Game = game ?? throw new Exception("Nessun plugin compatibile con questo file.");

                int pluginVersion = input.ReadInt32();
                if (pluginVersion > game.Version)
                {
                    throw new Exception("La versione installata del plugin non corrisponde a quella della traduzione.");
                }

                if (pluginVersion < game.Version)
                {
                    searchNewFiles = true;
                }

                string installPath = input.ReadString();
                if (!Directory.Exists(installPath))
                {
                    throw new Exception($"Non e' stata trovata la cartella di installazione: {installPath}");
                }

                result.InstallationPath = installPath;
                result.WorkPath = Path.GetDirectoryName(path);

                int containersCount = input.ReadInt32();
                for (var i = 0; i < containersCount; i++)
                {
                    string containerId = input.ReadString();
                    string containerPath = input.ReadString();
                    var containerType = (ContainerType)input.ReadInt32();
                    var container = new TranslationFileContainer(containerId, containerPath, containerType);

                    int fileCount = input.ReadInt32();
                    for (var j = 0; j < fileCount; j++)
                    {
                        string fileId = input.ReadString();
                        string filePath = input.ReadString();
                        string fileRelativePath = input.ReadString();
                        string fileName = input.ReadString();
                        string typeString = input.ReadString();

                        Type type = GetType(typeString, types);

                        TranslationFile file;

                        ConstructorInfo constructorInfo =
                            type.GetConstructor(new[] {typeof(string), typeof(string), typeof(string), typeof(Encoding)});

                        if (constructorInfo != null)
                        {
                            file = (TranslationFile)constructorInfo.Invoke(new object[]{game.Name, filePath, result.ChangesFolder, result.Game.FileEncoding});
                        }
                        else
                        {
                            constructorInfo = type.GetConstructor(new[] {typeof(string), typeof(string), typeof(string)});
                            if (constructorInfo != null)
                            {
                                file = (TranslationFile)constructorInfo.Invoke(new object[]{game.Name, filePath, result.ChangesFolder});
                            }
                            else
                            {
                                file = new TranslationFile(game.Name, filePath, result.ChangesFolder);
                            }
                        }

                        file.Id = fileId;
                        file.Name = fileName;
                        file.RelativePath = fileRelativePath;
                        container.AddFile(file);
                    }

                    result.FileContainers.Add(container);
                }

                if (searchNewFiles)
                {
                    UpdateTranslationFiles(result, worker);
                }

                return result;
            }
        }

        private static Type GetType(string typeName, Dictionary<string, Type> types)
        {
            if (types.ContainsKey(typeName))
            {
                return types[typeName];
            }

            Type type = Type.GetType(typeName);

            if (type != null)
            {
                types.Add(typeName, type);
                return type;
            }

            foreach (Assembly a in AppDomain.CurrentDomain.GetAssemblies())
            {
                type = a.GetType(typeName);
                if (type != null)
                {
                    types.Add(typeName, type);
                    return type;
                }
            }

            return null;
        }

        public void Export(IList<TranslationFileContainer> containers, ExportOptions options, BackgroundWorker worker)
        {
            PathHelper.DeleteDirectory(TempFolder);

            foreach (TranslationFileContainer container in containers)
            {
                if (worker.CancellationPending)
                {
                    worker.ReportProgress(0, "ANNULLATO");
                    throw new UserCancelException();
                }

                worker.ReportProgress(0, $"Ispezionando {container.Path}...");

                if (container.Type == ContainerType.Folder)
                {
                    string outputFolder = Path.GetFullPath(Path.Combine(ExportFolder, container.Path));

                    foreach (TranslationFile translationFile in container.Files)
                    {
                        if (translationFile.HasChanges || options.ForceRebuild)
                        {
                            translationFile.Rebuild(outputFolder);
                        }
                        else
                        {
                            if (!Game.ExportOnlyModifiedFiles)
                            {
                                string outputFile = Path.Combine(outputFolder, translationFile.RelativePath);
                                string dir = Path.GetDirectoryName(outputFile);
                                Directory.CreateDirectory(dir);

                                File.Copy(translationFile.Path, outputFile, true);
                            }
                        }
                    }
                }
                else
                {
                    string outputFile = Path.GetFullPath(Path.Combine(ExportFolder, container.Path));
                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }

                    worker.ReportProgress(0, "Copia in corso...");
                    // 1. Copiar todos los ficheros del contenedor a una carpeta temporal
                    string source = Path.Combine(ContainersFolder, container.Id);
                    string dest = Path.Combine(TempFolder, container.Id);
                    Directory.CreateDirectory(dest);

                    PathHelper.CloneDirectory(source, dest);

                    // 2. Crear los ficheros traducidos en esa carpeta temporal
                    worker.ReportProgress(0, "Generazione file temporanei...");
                    foreach (TranslationFile translationFile in container.Files)
                    {
                        if (translationFile.HasChanges || options.ForceRebuild)
                        {
                            translationFile.Rebuild(dest);
                        }
                        else
                        {
                            if (Game.ExportOnlyModifiedFiles)
                            {
                                var file = Path.Combine(dest, translationFile.RelativePath);
                                File.Delete(file);
                            }
                        }
                    }

                    // 3. Empaquetar
                    worker.ReportProgress(0, "Compressione file...");
                    Game.RepackFile(dest, outputFile, options.UseCompression);

                    // 4. Eliminar la carpeta temporal
                    if (!options.SaveTempFiles)
                    {
                        try
                        {
                            Directory.Delete(dest, true);
                        }
                        catch (IOException)
                        {
                            Thread.Sleep(0);
                            Directory.Delete(dest, true);
                        }
                    }
                }
            }
        }

        public IList<Tuple<TranslationFileContainer, TranslationFile>> SearchInFiles(string searchString, BackgroundWorker worker)
        {
            var result = new ConcurrentBag<Tuple<TranslationFileContainer, TranslationFile>>();

            foreach (TranslationFileContainer container in FileContainers)
            {
                if (worker.CancellationPending)
                {
                    worker.ReportProgress(0, "ANNULLATO");
                    throw new UserCancelException();
                }

                worker.ReportProgress(0, $"Ispezionando {container.Path}...");

                //foreach (var file in container.Files)
                Parallel.ForEach(container.Files, file =>
                {
                    bool found = file.Search(searchString);

                    if (found)
                    {
                        result.Add(new Tuple<TranslationFileContainer, TranslationFile>(container, file));
                    }
                });

            }

            return result.ToList();
        }

        private static void UpdateTranslationFiles(TranslationProject project, BackgroundWorker worker)
        {
            GameFileContainer[] containers = project.Game.GetContainers(project.InstallationPath);
            foreach (GameFileContainer container in containers)
            {
                if (worker.CancellationPending)
                {
                    worker.ReportProgress(0, "ANNULLATO");
                    throw new UserCancelException();
                }

                TranslationFileContainer translationContainer =
                    project.FileContainers.FirstOrDefault(x => x.Path == container.Path && x.Type == container.Type);

                var addNewContainer = false;
                var addedFiles = 0;

                if (translationContainer == null)
                {
                    translationContainer = new TranslationFileContainer(container.Path, container.Type);
                    addNewContainer = true;
                }

                string extractionContainerPath = Path.Combine(project.ContainersFolder, translationContainer.Id);
                Directory.CreateDirectory(extractionContainerPath);

                string containerPath = Path.GetFullPath(Path.Combine(project.InstallationPath, container.Path));

                worker.ReportProgress(0, $"Ispezionando {container.Path}...");
                if (container.Type == ContainerType.CompressedFile)
                {
                    if (File.Exists(containerPath))
                    {
                        if (addNewContainer)
                        {
                            project.Game.PreprocessContainer(translationContainer, containerPath, extractionContainerPath);
                            project.Game.ExtractFile(containerPath, extractionContainerPath);
                        }

                        foreach (GameFileSearch fileSearch in container.FileSearches)
                        {
                            worker.ReportProgress(0, $"Cercando {fileSearch.RelativePath}\\{fileSearch.SearchPattern}...");
                            string[] foundFiles = fileSearch.GetFiles(extractionContainerPath);
#if DEBUG
                            foreach (string f in foundFiles)
#else
                            Parallel.ForEach(foundFiles, f =>
#endif
                            {
                                string relativePath =
                                    PathHelper.GetRelativePath(extractionContainerPath, Path.GetFullPath(f));
                                Type type = fileSearch.FileType;

                                TranslationFile translationFile =
                                    translationContainer.Files.FirstOrDefault(x => x.RelativePath == relativePath);

                                if (translationFile == null)
                                {
                                    ConstructorInfo constructorInfo =
                                        type.GetConstructor(new[]
                                            {typeof(string), typeof(string), typeof(string), typeof(Encoding)});

                                    if (constructorInfo != null)
                                    {
                                        translationFile = (TranslationFile) constructorInfo.Invoke(new object[]
                                            {project.Game.Name, f, project.ChangesFolder, project.Game.FileEncoding});
                                    }
                                    else
                                    {
                                        constructorInfo = type.GetConstructor(new[]
                                            {typeof(string), typeof(string), typeof(string)});
                                        if (constructorInfo != null)
                                        {
                                            translationFile = (TranslationFile) constructorInfo.Invoke(new object[]
                                                {project.Game.Name, f, project.ChangesFolder});
                                        }
                                        else
                                        {
                                            translationFile = new TranslationFile(project.Game.Name, f,
                                                project.ChangesFolder);
                                        }
                                    }

                                    translationFile.RelativePath = relativePath;

                                    if (translationFile.Type == FileType.TextFile)
                                    {
                                        if (translationFile.SubtitleCount > 0)
                                        {
                                            translationContainer.AddFile(translationFile);
                                            addedFiles++;
                                        }
                                    }
                                    else
                                    {
                                        translationContainer.AddFile(translationFile);
                                        addedFiles++;
                                    }
                                }
#if DEBUG
                            }
#else
                            });
#endif
                        }

                        project.Game.PostprocessContainer(translationContainer, containerPath, extractionContainerPath);

                        worker.ReportProgress(0, $"{addedFiles} file trovati e aggiunti");
                    }
                    else
                    {
                        worker.ReportProgress(0, $"ERROR: Il file compresso non esiste: {containerPath}");
                        continue;
                    }
                }
                else
                {
                    project.Game.PreprocessContainer(translationContainer, containerPath, extractionContainerPath);
                    foreach (GameFileSearch fileSearch in container.FileSearches)
                    {
                        worker.ReportProgress(0, $"Cercando {fileSearch.RelativePath}\\{fileSearch.SearchPattern}...");
                        string[] foundFiles = fileSearch.GetFiles(containerPath);

#if DEBUG
                        foreach (string f in foundFiles)
#else
                        Parallel.ForEach(foundFiles, f =>
#endif
                        {
                            string relativePath = PathHelper.GetRelativePath(containerPath, Path.GetFullPath(f));

                            string destinationFileName =
                                Path.GetFullPath(Path.Combine(extractionContainerPath, relativePath));
                            string destPath = Path.GetDirectoryName(destinationFileName);
                            Directory.CreateDirectory(destPath);

                            if (!File.Exists(destinationFileName))
                            {
                                File.Copy(f, destinationFileName);
                            }

                            Type type = fileSearch.FileType;

                            TranslationFile translationFile =
                                translationContainer.Files.FirstOrDefault(x => x.RelativePath == relativePath);

                            if (translationFile == null)
                            {
                                ConstructorInfo constructorInfo =
                                    type.GetConstructor(new[] {typeof(string), typeof(string), typeof(string), typeof(Encoding)});

                                if (constructorInfo != null)
                                {
                                    translationFile = (TranslationFile)constructorInfo.Invoke(new object[]{project.Game.Name, destinationFileName, project.ChangesFolder, project.Game.FileEncoding});
                                }
                                else
                                {
                                    constructorInfo = type.GetConstructor(new[] {typeof(string), typeof(string), typeof(string)});
                                    if (constructorInfo != null)
                                    {
                                        translationFile = (TranslationFile)constructorInfo.Invoke(new object[]{project.Game.Name, destinationFileName, project.ChangesFolder});
                                    }
                                    else
                                    {
                                        translationFile = new TranslationFile(project.Game.Name, destinationFileName, project.ChangesFolder);
                                    }
                                }

                                translationFile.RelativePath = relativePath;

                                if (translationFile.Type == FileType.TextFile)
                                {
                                    if (translationFile.SubtitleCount > 0)
                                    {
                                        translationContainer.AddFile(translationFile);
                                        addedFiles++;
                                    }
                                }
                                else
                                {
                                    translationContainer.AddFile(translationFile);
                                    addedFiles++;
                                }
                            }
#if DEBUG
                        }
#else
                        });
#endif

                        project.Game.PostprocessContainer(translationContainer, containerPath, extractionContainerPath);
                        worker.ReportProgress(0, $"{addedFiles} file trovati e aggiunti");
                    }
                }

                if (addNewContainer && translationContainer.Files.Count > 0)
                {
                    project.FileContainers.Add(translationContainer);
                }
            }
        }

        public void ExportPo(string path, BackgroundWorker worker)
        {
            foreach (TranslationFileContainer container in FileContainers)
            {
                if (worker.CancellationPending)
                {
                    worker.ReportProgress(0, "ANNULLATO");
                    throw new UserCancelException();
                }

                worker.ReportProgress(0, $"Ispezionando {container.Path}...");

                foreach (TranslationFile file in container.Files)
                {
                    string filePath = Path.Combine(path, container.Path, file.RelativePath);
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string outputPath = string.Concat(@"\\?\", Path.GetFullPath(Path.Combine(Path.GetDirectoryName(filePath), string.Concat(fileName, ".po"))));
                    file.ExportPo(outputPath);
                }
            }
        }

        public void ExportImages(string path, BackgroundWorker worker)
        {
            foreach (TranslationFileContainer container in FileContainers)
            {
                if (worker.CancellationPending)
                {
                    worker.ReportProgress(0, "ANNULLATO");
                    throw new UserCancelException();
                }

                worker.ReportProgress(0, $"Ispezionando {container.Path}...");

                foreach (TranslationFile file in container.Files)
                {
                    string filePath = Path.Combine(path, container.Path, file.RelativePath);
                    string fileName = file.GetExportFilename();
                    string outputPath = string.Concat(@"\\?\", Path.GetFullPath(Path.Combine(Path.GetDirectoryName(filePath), fileName)));
                    file.ExportImage(outputPath);
                }
            }
        }

        public void ExportXlsx(string path, BackgroundWorker worker)
        {
            foreach (TranslationFileContainer container in FileContainers)
            {
				if (worker.CancellationPending)
				{
					worker.ReportProgress(0, "ANNULLATO");
                    throw new UserCancelException();
                }

				

				foreach (TranslationFile file in container.Files)
				{
					string filePath = Path.Combine(path, container.Path, file.RelativePath);
					string fileName = Path.GetFileNameWithoutExtension(filePath);
					string outputPath = string.Concat(@"\\?\", Path.GetFullPath(Path.Combine(Path.GetDirectoryName(filePath), string.Concat(fileName, ".xlsx"))));

                    if(fileName.Contains(".tex"))
                    {
                        continue;
                    }

                    if(!filePath.Contains("luabytecode"))
                    {
                        continue;
                    }

                    worker.ReportProgress(-1, "Nome: " + fileName);

                    // TODO
                    var btf = (BinaryTextFile)file;
                    btf._view = new GridView(btf);
					btf._subtitles = btf.GetSubtitles();
					btf._view.LoadData(btf._subtitles.Where(x => !string.IsNullOrEmpty(x.Text)).ToList());

                    btf._view.ExportExcel();
				}
			}
        }

        public void ImportPoFromDirectory(string directory, string fileName, TranslationFile file, BackgroundWorker worker)
        {
            string newFilename = string.Concat(fileName, ".po");
			string[] files = Directory.GetFiles(directory,
								newFilename, SearchOption.AllDirectories);

			if (files.Length > 0)
			{
                worker.ReportProgress(-1, "ATTENZIONE: " + files.Length.ToString() + " File trovati!");
				foreach (string s in files)
				{
                    worker.ReportProgress(-1, "File trovato: " + files[files.Length - 1].ToString());
					file.ImportPo(s);
				}
			}
			else
			{
				//worker.ReportProgress(-1, "ERRORE: Nessun file trovato in: " + directory + " chiamato " + newFilename);
			}
		}
        
        public void ImportPo(string path, BackgroundWorker worker)
        {
            foreach (TranslationFileContainer container in FileContainers)
            {
                if (worker.CancellationPending)
                {
                    worker.ReportProgress(0, "ANNULLATO");
                    throw new UserCancelException();
                }

                worker.ReportProgress(0, $"Ispezionando {container.Path}...");

                foreach (TranslationFile file in container.Files)
                {
                    string filePath = Path.Combine(path, container.Path, file.RelativePath);
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string inputPath = string.Concat(@"\\?\", Path.GetFullPath(Path.Combine(Path.GetDirectoryName(filePath), string.Concat(fileName, ".po"))));

                    if(fileName.Contains(".tex"))
                    {
                        continue;
					}

					if (fileName.Contains(".ttf"))
					{
						continue;
					}

					if (!File.Exists(inputPath))
                    {
                        filePath = Path.Combine(path, file.RelativePath);
                        inputPath = string.Concat(@"\\?\", Path.GetFullPath(Path.Combine(Path.GetDirectoryName(filePath), string.Concat(fileName, ".po"))));
                    }

                    if (File.Exists(inputPath))
                    {
                        file.ImportPo(inputPath);
                    }
                    else
                    {
                        // Comprobamos si el fichero está partido
                        string directory = string.Concat(@"\\?\", Path.GetFullPath(Path.GetDirectoryName(Path.Combine(path, container.Path, file.RelativePath))));
                        if (Directory.Exists(directory))
                        {
                            ImportPoFromDirectory(directory, fileName, file, worker);
                        } else
                        {
							//worker.ReportProgress(-1, "ERRORE: La directory " + directory + " non esiste!");

                            directory = string.Concat(@"\\?\", Path.GetFullPath(path));
							ImportPoFromDirectory(directory, fileName, file, worker);
						}
                    }
                }
            }
        }

        public void ImportImages(string path, BackgroundWorker worker)
        {
            foreach (TranslationFileContainer container in FileContainers)
            {
                if (worker.CancellationPending)
                {
                    worker.ReportProgress(0, "ANNULLATO");
                    throw new UserCancelException();
                }

                worker.ReportProgress(0, $"Ispezionando {container.Path}...");

                foreach (TranslationFile file in container.Files)
                {
                    string filePath = Path.Combine(path, container.Path, file.RelativePath);
                    string fileName = file.GetExportFilename();
                    string inputPath = string.Concat(@"\\?\", Path.GetFullPath(Path.Combine(Path.GetDirectoryName(filePath), fileName)));

                    if (File.Exists(inputPath))
                    {
                        file.ImportImage(inputPath);
                    }
                }
            }
        }
    }
}
