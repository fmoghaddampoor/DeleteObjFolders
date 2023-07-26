namespace CloseAllTabs
{
    using System;
    using System.Diagnostics;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell.Events;

    public class DeleteIISExpressTraceLogFilesFolder : DeleteBase
    {
        private DeleteIISExpressTraceLogFilesFolder(DTE2 dte, Options options)
        {
            this._dte = dte;
            this._options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static DeleteIISExpressTraceLogFilesFolder Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new DeleteIISExpressTraceLogFilesFolder(dte, options);
        }

        private void Execute()
        {
            if (!this._options.DeleteIISExpressTraceLogFilesFolder)
            {
                return;
            }

            try
            {
                var root = GetIISExpressTraceLogFilesFolder();
                this.DeleteFiles(root);
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
            }
        }
    }
}