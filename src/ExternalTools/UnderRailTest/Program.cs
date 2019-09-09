﻿using System;
using System.IO;
using UnderRailLib;
using UnderRailLib.AssemblyResolver;

namespace UnderRailTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var assemblyResolver = new AssemblyResolver();
            assemblyResolver.Initialize();
            var eventHandler = new ResolveEventHandler(assemblyResolver.ResolveAssembly);
            AppDomain.CurrentDomain.AssemblyResolve += eventHandler;
            Binder.SetAssemblyResolver(assemblyResolver);

            var dialogManager = new DialogManager();

            var files = Directory.EnumerateFiles(@"H:\Games\Underrail\data\dialogs", "*.udlg",
                SearchOption.AllDirectories);

            foreach (var file in files)
            {
                var model = dialogManager.LoadModel(file);

                if (model != null)
                {
                    var output = file.Replace("dialogs", "dialogs2");
                    dialogManager.SaveModel(model, output);
                }
            }
            
        }
    }
}