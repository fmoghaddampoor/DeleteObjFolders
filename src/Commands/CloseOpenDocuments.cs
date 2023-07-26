namespace CloseAllTabs
{
    using System;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using SolutionEvents = Microsoft.VisualStudio.Shell.Events.SolutionEvents;

    public class CloseOpenDocuments
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly DTE2 _dte;
        private readonly Options _options;

        private CloseOpenDocuments(IServiceProvider serviceProvider, DTE2 dte, Options options)
        {
            this._serviceProvider = serviceProvider;
            this._dte = dte;
            this._options = options;
        }

        public static CloseOpenDocuments Instance { get; private set; }

        public static void Initialize(Package serviceProvider, DTE2 dte, Options options)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Instance = new CloseOpenDocuments(serviceProvider, dte, options);
            SolutionEvents.OnBeforeCloseSolution += (s, e) => Instance.Execute();
        }

        private void Execute()
        {
            if (!this._options.CloseDocuments)
            {
                return;
            }

            foreach (Document document in this._dte.Documents)
            {
                var filePath = document.FullName;

                // Don't close pinned files
                if (VsShellUtilities.IsDocumentOpen(
                        this._serviceProvider, filePath, VSConstants.LOGVIEWID_Primary, out var hierarchy, out var itemId,
                        out var frame))
                {
                    ErrorHandler.ThrowOnFailure(frame.GetProperty((int)__VSFPROPID5.VSFPROPID_IsPinned, out var propVal));

                    if (bool.TryParse(propVal.ToString(), out var isPinned) && !isPinned)
                    {
                        document.Close();
                    }
                }
                else
                {
                    document.Close();
                }
            }
        }
    }
}