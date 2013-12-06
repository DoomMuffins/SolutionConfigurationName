using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;

namespace SolutionConfigurationName
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidSolutionConfigurationNamePkgString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
    public sealed class MainSite : Package
    {
        private static IVsSolutionBuildManager _iVsSolutionBuildManager;

        private uint _solutionEventsCookie;
        // Because we load on UICONTEXT.NoSolution_string, we make sure that UpdateSolutionEvents' default 
        // constructor runs and the vars are first set as early as possible
        private readonly UpdateSolutionEvents _updateSolutionEvents = new UpdateSolutionEvents();

        public static DTE DTE { get; private set; }

        public MainSite() { }

        protected override void Initialize()
        {
            base.Initialize();

            InitializeGlobalServices();

            uint pdwCookie;
            int hr = _iVsSolutionBuildManager.AdviseUpdateSolutionEvents(_updateSolutionEvents, out pdwCookie);
            ErrorHandler.ThrowOnFailure(hr);
            _solutionEventsCookie = pdwCookie;
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (_solutionEventsCookie != default(uint))
            {
                int hr = _iVsSolutionBuildManager.UnadviseUpdateSolutionEvents(_solutionEventsCookie);
                ErrorHandler.ThrowOnFailure(hr);
            }
        }

        private void InitializeGlobalServices()
        {
            if (DTE == null)
            {
                DTE = GetGlobalService(typeof (DTE)) as DTE;
            }
            if (_iVsSolutionBuildManager == null)
            {
                _iVsSolutionBuildManager = GetGlobalService(typeof (SVsSolutionBuildManager)) as IVsSolutionBuildManager;
            }
        }
    }
}
