using task = System.Threading.Tasks.Task;

namespace CloseAllTabs
{
    using System;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Commands;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;

    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)]
    [ProvideAutoLoad(UIContextGuids.SolutionHasSingleProject, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(UIContextGuids.SolutionHasMultipleProjects, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideOptionPage(
        typeof(Options), "Environment", Vsix.Name, 101, 102,
        true, new string[] { }, ProvidesLocalizedCategoryName = false)]
    [Guid(Vsix.Id)]
    public sealed class CleanOnClosePackage : AsyncPackage
    {
        protected override async task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync();
            var dte = await this.GetServiceAsync(typeof(DTE)) as DTE2;

            var options = (Options)this.GetDialogPage(typeof(Options));

            CloseOpenDocuments.Initialize(this, dte, options);
            CollapseFolders.Initialize(dte, options);
            SolutionExplorerFocus.Initialize(dte, options);
            DeleteBinFolder.Initialize(dte, options);
            DeleteObjFolder.Initialize(dte, options);
            DeleteTestResultsFolder.Initialize(dte, options);
            DeleteDotVsFolder.Initialize(dte, options);
            DeleteIisExpressLogsFolder.Initialize(dte, options);
            DeleteIisExpressTraceLogFilesFolder.Initialize(dte, options);
        }
    }
}