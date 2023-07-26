// ReSharper disable All
namespace CloseAllTabs.Commands
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell.Events;

    public class DeleteObjFolder : DeleteBase
    {
        private DeleteObjFolder(DTE2 dte, Options options)
        {
            this.Dte = dte;
            this.Options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static DeleteObjFolder Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new DeleteObjFolder(dte, options);
        }

        private void Execute()
        {
            if (!this.Options.DeleteObjFolder)
            {
                return;
            }

            try
            {
                foreach (var project in this.GetAllProjects())
                {
                    var root = GetProjectRootFolder(project);

                    if (root == null)
                    {
                        return;
                    }

                    var obj = Path.Combine(root, "obj");


                    this.DeleteFiles(obj);
                }
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
            }
        }
    }
}