// ReSharper disable All
namespace CloseAllTabs.Commands
{
    using System;
    using System.Diagnostics;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell.Events;

    public class DeleteIisExpressTraceLogFilesFolder : DeleteBase
    {
        private DeleteIisExpressTraceLogFilesFolder(DTE2 dte, Options options)
        {
            this.Dte = dte;
            this.Options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static DeleteIisExpressTraceLogFilesFolder Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new DeleteIisExpressTraceLogFilesFolder(dte, options);
        }

        private void Execute()
        {
            if (!this.Options.DeleteIisExpressTraceLogFilesFolder)
            {
                return;
            }

            try
            {
                var root = GetIisExpressTraceLogFilesFolder();
                this.DeleteFiles(root);
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
            }
        }
    }
}