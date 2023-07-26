// ReSharper disable All
namespace CloseAllTabs.Commands
{
    using EnvDTE;
    using EnvDTE80;
    using SolutionEvents = Microsoft.VisualStudio.Shell.Events.SolutionEvents;

    public class SolutionExplorerFocus
    {
        private readonly DTE2 dte;
        private readonly Options options;

        private SolutionExplorerFocus(DTE2 dte, Options options)
        {
            this.dte = dte;
            this.options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static SolutionExplorerFocus Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new SolutionExplorerFocus(dte, options);
        }

        private void Execute()
        {
            if (!this.options.FocusSolutionExplorer)
            {
                return;
            }

            var solExp = this.dte.Windows.Item(Constants.vsWindowKindSolutionExplorer);

            if (solExp != null)
            {
                solExp.Activate();
            }
        }
    }
}