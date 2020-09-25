using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace VSIXTPLDataFlowDebuggerVisualizer
{
    /// <summary>
    ///     This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         The minimum requirement for a class to be considered a valid package for Visual Studio
    ///         is to implement the IVsPackage interface and register itself with the shell.
    ///         This package uses the helper classes defined inside the Managed Package Framework (MPF)
    ///         to do it: it derives from the Package class that provides the implementation of the
    ///         IVsPackage interface and uses the registration attributes defined in the framework to
    ///         register itself and its components with the shell. These attributes tell the pkgdef creation
    ///         utility what data to put into .pkgdef file.
    ///     </para>
    ///     <para>
    ///         To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...
    ///         &gt; in .vsixmanifest file.
    ///     </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly",
        Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class VsixPackage : AsyncPackage
    {
        /// <summary>
        ///     VsixPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "4d43737a-e8e5-4906-91fd-f08adecfcc7f";

        private readonly string[] VisualizerAssmNames =
        {
            "TPLDataFlowDebuggerVisualizer.dll", "GraphSharp.dll", "GraphSharp.Controls.dll", "QuickGraph.Data.dll",
            "QuickGraph.dll", "QuickGraph.Graphviz.dll", "QuickGraph.Serialization.dll", "WPFExtensions.dll"
        };

        #region Package Members

        protected override async Task InitializeAsync(CancellationToken cancellationToken,
            IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            try
            {
                // Get the destination folder for visualizers
                if (await GetServiceAsync(typeof(SVsShell)) is IVsShell shell)
                {
                    shell.GetProperty((int) __VSSPROPID2.VSSPROPID_VisualStudioDir,
                        out var documentsFolderFullNameObject);

                    VisualizerAssmNames.ToList().ForEach(s => CopyDll(s, documentsFolderFullNameObject.ToString()));
                }
            }
            catch (Exception ex)
            {
                // TODO: Handle exception
                Debug.WriteLine(ex.ToString());
            }
        }

        private void CopyDll(string fileName, string documentsFolderFullName)
        {
            // The Visualizer dll is in the same folder than the package because its project is added as reference to this project,
            // so it is included inside the .vsix file. We only need to deploy it to the correct destination folder.
            var sourceFolderFullName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var destinationFolderFullName = Path.Combine(documentsFolderFullName, "Visualizers");

            var sourceFileFullName = Path.Combine(sourceFolderFullName, fileName);
            var destinationFileFullName = Path.Combine(destinationFolderFullName, fileName);

            CopyFileIfNewerVersion(sourceFileFullName, destinationFileFullName);
        }

        private void CopyFileIfNewerVersion(string sourceFileFullName, string destinationFileFullName)
        {
            bool copy = false;

            if (File.Exists(destinationFileFullName))
            {
                var sourceFileVersionInfo = FileVersionInfo.GetVersionInfo(sourceFileFullName);
                var destinationFileVersionInfo = FileVersionInfo.GetVersionInfo(destinationFileFullName);
                if (sourceFileVersionInfo.FileMajorPart > destinationFileVersionInfo.FileMajorPart)
                    copy = true;
                else if (sourceFileVersionInfo.FileMajorPart == destinationFileVersionInfo.FileMajorPart
                         && sourceFileVersionInfo.FileMinorPart > destinationFileVersionInfo.FileMinorPart)
                    copy = true;
            }
            else
            {
                // First time
                copy = true;
            }

            if (copy) File.Copy(sourceFileFullName, destinationFileFullName, true);
        }

        #endregion
    }
}