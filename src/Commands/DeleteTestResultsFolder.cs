// ReSharper disable All
namespace CloseAllTabs.Commands
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
            this.Dte = dte;
            this.Options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static DeleteTestResultsFolder Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new DeleteTestResultsFolder(dte, options);
        }

        private void Execute()
        {
            if (!this.Options.DeleteTestResultsFolder)
            {
                return;
            }

            try
            {
                var root = GetSolutionRootFolder(this.Dte.Solution);
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