namespace CloseAllTabs
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell.Events;

    public class DeleteTestResultsFolder : DeleteBase
    {
        private DeleteTestResultsFolder(DTE2 dte, Options options)
        {
            this._dte = dte;
            this._options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static DeleteTestResultsFolder Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new DeleteTestResultsFolder(dte, options);
        }

        private void Execute()
        {
            if (!this._options.DeleteTestResultsFolder)
            {
                return;
            }

            try
            {
                var root = GetSolutionRootFolder(this._dte.Solution);
                var testResults = Path.Combine(root, "TestResults");
                this.DeleteFiles(testResults);
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
            }
        }
    }
}