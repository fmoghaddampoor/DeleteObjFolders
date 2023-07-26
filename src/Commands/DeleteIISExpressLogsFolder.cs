// ReSharper disable All
namespace CloseAllTabs.Commands
{
    using System;
    using System.Diagnostics;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell.Events;

    public class DeleteIisExpressLogsFolder : DeleteBase
    {
        private DeleteIisExpressLogsFolder(DTE2 dte, Options options)
        {
            this.Dte = dte;
            this.Options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static DeleteIisExpressLogsFolder Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new DeleteIisExpressLogsFolder(dte, options);
        }

        private void Execute()
        {
            if (!this.Options.DeleteIisExpressLogsFolder)
            {
                return;
            }

            try
            {
                var root = GetIisExpressLogsFolder();
                this.DeleteFiles(root);
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
            }
        }
    }
}