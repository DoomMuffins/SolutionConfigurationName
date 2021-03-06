﻿using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using EnvDTE80;
using Microsoft.Build.Evaluation;

namespace SolutionConfigurationName
{
    public class UpdateSolutionEvents : IVsUpdateSolutionEvents
    {
        private const string SlnConfigPropertyName = "SolutionConfigurationName";
        private const string SlnPlatformPropertyName = "SolutionPlatformName";

        private string _lastConfigurationName = string.Empty;
        private string _lastPlatformName = string.Empty;

        public UpdateSolutionEvents()
        {
            // defining these / making sure they exist early prevents a race between Visual Studio reading OutputPath
            // and OnActiveProjectCfgChange below defining the global properties (if they are undefined, VS permanently
            // ignores them
            ProjectCollection global = ProjectCollection.GlobalProjectCollection;
            global.SetGlobalProperty(SlnConfigPropertyName, string.Empty);
            global.SetGlobalProperty(SlnPlatformPropertyName, string.Empty);
        }

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
                global.SetGlobalProperty(SlnConfigPropertyName, currentConfigurationName);
                _lastConfigurationName = currentConfigurationName;
            }

            if (currentPlatformName != _lastPlatformName)
            {
                global.SetGlobalProperty(SlnPlatformPropertyName, currentPlatformName);
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
