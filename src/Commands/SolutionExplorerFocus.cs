namespace CloseAllTabs
{
    using EnvDTE;
    using EnvDTE80;
    using SolutionEvents = Microsoft.VisualStudio.Shell.Events.SolutionEvents;

    public class SolutionExplorerFocus
    {
        private readonly DTE2 _dte;
        private readonly Options _options;

        private SolutionExplorerFocus(DTE2 dte, Options options)
        {
            this._dte = dte;
            this._options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static SolutionExplorerFocus Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new SolutionExplorerFocus(dte, options);
        }

        private void Execute()
        {
            if (!this._options.FocusSolutionExplorer)
            {
                return;
            }

            var solExp = this._dte.Windows.Item(Constants.vsWindowKindSolutionExplorer);

            if (solExp != null)
            {
                solExp.Activate();
            }
        }
    }
}