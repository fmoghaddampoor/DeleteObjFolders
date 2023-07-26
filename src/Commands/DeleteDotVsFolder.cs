namespace CloseAllTabs
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell.Events;

    public class DeleteDotVsFolder : DeleteBase
    {
        private DeleteDotVsFolder(DTE2 dte, Options options)
        {
            this._dte = dte;
            this._options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static DeleteDotVsFolder Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new DeleteDotVsFolder(dte, options);
        }

        private void Execute()
        {
            if (!this._options.DeleteDotVsFolder)
            {
                return;
            }

            try
            {
                var root = GetSolutionRootFolder(this._dte.Solution);
                var dotVs = Path.Combine(root, ".vs");
                this.DeleteFiles(dotVs);
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
            }
        }
    }
}