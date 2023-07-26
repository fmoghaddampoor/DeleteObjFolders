namespace CloseAllTabs
{
    using System;
    using System.Diagnostics;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell.Events;

    public class DeleteIISExpressLogsFolder : DeleteBase
    {
        private DeleteIISExpressLogsFolder(DTE2 dte, Options options)
        {
            this._dte = dte;
            this._options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static DeleteIISExpressLogsFolder Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new DeleteIISExpressLogsFolder(dte, options);
        }

        private void Execute()
        {
            if (!this._options.DeleteIISExpressLogsFolder)
            {
                return;
            }

            try
            {
                var root = GetIISExpressLogsFolder();
                this.DeleteFiles(root);
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
            }
        }
    }
}