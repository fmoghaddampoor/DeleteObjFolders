// ReSharper disable All
namespace CloseAllTabs.Commands
{
    using System.Linq;
    using EnvDTE;
    using EnvDTE80;
    using SolutionEvents = Microsoft.VisualStudio.Shell.Events.SolutionEvents;

    public class CollapseFolders
    {
        private readonly DTE2 dte;
        private readonly Options options;

        private CollapseFolders(DTE2 dte, Options options)
        {
            this.dte = dte;
            this.options = options;

            SolutionEvents.OnBeforeCloseSolution += (s, e) => this.Execute();
        }

        public static CollapseFolders Instance { get; private set; }

        public static void Initialize(DTE2 dte, Options options)
        {
            Instance = new CollapseFolders(dte, options);
        }

        private void Execute()
        {
            if (!this.options.CollapseOn)
            {
                return;
            }

            var hierarchy = this.dte.ToolWindows.SolutionExplorer.UIHierarchyItems;

            try
            {
                this.dte.SuppressUI = true;
                this.CollapseHierarchy(hierarchy);
            }
            finally
            {
                this.dte.SuppressUI = false;
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
            if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder && this.options.CollapseSolutionFolders)
            {
                return true;
            }

            // Collapse projects if enabled in settings
            if (project.Kind != ProjectKinds.vsProjectKindSolutionFolder && this.options.CollapseProjects)
            {
                return true;
            }

            return false;
        }
    }
}