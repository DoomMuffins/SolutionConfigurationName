using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using EnvDTE80;
using Microsoft.Build.Evaluation;

namespace SolutionConfigurationName
{
    public class UpdateSolutionEvents : IVsUpdateSolutionEvents
    {
        private string _lastConfigurationName = string.Empty;
        private string _lastPlatformName = string.Empty;

        /// <summary>
        /// This is called every time the active configuration of a project changes (according to MSDN).
        /// It also seems to be called once when the solution configuration itself is changed (even if
        /// it doesn't cause the active configuration of a project to change) - empirically, for a solution
        /// with one project this will be called TWICE. I therefore use it as an indicator for solution
        /// configuration changes because its granularity seems exactly right (and it doesn't wait for a 
        /// build to be called).
        /// </summary>
        public int OnActiveProjectCfgChange(IVsHierarchy pIVsHierarchy)
        {
            ProjectCollection global = ProjectCollection.GlobalProjectCollection;
            SolutionConfiguration2 configuration =
                (SolutionConfiguration2)MainSite.DTE.Solution.SolutionBuild.ActiveConfiguration;
            string currentConfigurationName = configuration.Name;
            string currentPlatformName = configuration.PlatformName;
            
            if (currentConfigurationName != _lastConfigurationName)
            {
                global.SetGlobalProperty("SolutionConfigurationName", currentConfigurationName);
                _lastConfigurationName = currentConfigurationName;
            }

            if (currentPlatformName != _lastPlatformName)
            {
                global.SetGlobalProperty("SolutionPlatformName", currentPlatformName);
                _lastPlatformName = currentConfigurationName;
            }

            return VSConstants.S_OK;
        }

        public int UpdateSolution_Begin(ref int pfCancelUpdate)
        {
            return VSConstants.S_OK;
        }

        public int UpdateSolution_Cancel()
        {
            return VSConstants.S_OK;
        }

        public int UpdateSolution_Done(int fSucceeded, int fModified, int fCancelCommand)
        {
            return VSConstants.S_OK;
        }

        public int UpdateSolution_StartUpdate(ref int pfCancelUpdate)
        {
            return VSConstants.S_OK;
        }
    }
}
