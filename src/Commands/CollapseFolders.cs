namespace CloseAllTabs
{
    using System.Linq;
    using EnvDTE;
    using EnvDTE80;
    using SolutionEvents = Microsoft.VisualStudio.Shell.Events.SolutionEvents;

    public class CollapseFolders
    {
        private readonly DTE2 _dte;
        private readonly Options _options;

        private CollapseFolders(DTE2 dte, Options options)
        {
            this._dte = dte;
            this._options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static CollapseFolders Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new CollapseFolders(dte, options);
        }

        private void Execute()
        {
            if (!this._options.CollapseOn)
            {
                return;
            }

            var hierarchy = this._dte.ToolWindows.SolutionExplorer.UIHierarchyItems;

            try
            {
                this._dte.SuppressUI = true;
                this.CollapseHierarchy(hierarchy);
            }
            finally
            {
                this._dte.SuppressUI = false;
            }
        }

        private void CollapseHierarchy(UIHierarchyItems hierarchy)
        {
            foreach (var item in hierarchy.Cast<UIHierarchyItem>().Where(item => item.UIHierarchyItems.Count > 0))
            {
                this.CollapseHierarchy(item.UIHierarchyItems);

                if (this.ShouldCollapse(item))
                {
                    item.UIHierarchyItems.Expanded = false;
                }
            }
        }

        private bool ShouldCollapse(UIHierarchyItem item)
        {
            if (!item.UIHierarchyItems.Expanded)
            {
                return false;
            }

            // Always collapse files and folders
            if (!(item.Object is Project project))
            {
                return true;
            }

            // Collapse solution folders if enabled in settings
            if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder && this._options.CollapseSolutionFolders)
            {
                return true;
            }

            // Collapse projects if enabled in settings
            if (project.Kind != ProjectKinds.vsProjectKindSolutionFolder && this._options.CollapseProjects)
            {
                return true;
            }

            return false;
        }
    }
}