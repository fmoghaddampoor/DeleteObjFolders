namespace CloseAllTabs.Commands
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell.Events;

    public class DeleteBinFolder : DeleteBase
    {
        private DeleteBinFolder(DTE2 dte, Options options)
        {
            this.Dte = dte;
            this.Options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static DeleteBinFolder Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new DeleteBinFolder(dte, options);
        }

        private void Execute()
        {
            if (!this.Options.DeleteBinFolder)
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

                    var bin = Path.Combine(root, "bin");
                    var obj = Path.Combine(root, "obj");

                    this.DeleteFiles(bin, obj);
                }
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
            }
        }
    }
}